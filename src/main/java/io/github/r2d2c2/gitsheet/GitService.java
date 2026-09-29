package io.github.r2d2c2.gitsheet;

import java.io.*;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;

/** Git work is serialized; literal pathspecs protect document boundaries. */
public final class GitService {
    private GitService() {}
    public record Result(int code, String output) {}
    public enum Provider {
        LOCAL("로컬 Git만", ""),
        GITHUB("GitHub", "https://github.com/owner/repository.git"),
        GITLAB("GitLab", "https://gitlab.com/group/repository.git"),
        GITEA("Gitea", "https://your-gitea.example/owner/repository.git"),
        OTHER("기타 Git 서버", "https://your-server.example/team/repository.git");
        private final String label, example;
        Provider(String label, String example) { this.label = label; this.example = example; }
        public String example() { return example; }
        public String remoteName() { return "gitsheet-" + name().toLowerCase(Locale.ROOT); }
        @Override public String toString() { return label; }
    }
    public record Settings(Path root, Provider provider, String url, String name, String email, String branch) {}
    public static Result run(Path directory, String... arguments) throws IOException, InterruptedException {
        var command = new ArrayList<>(List.of("git", "--literal-pathspecs", "-c", "core.quotepath=false", "-c", "color.ui=false", "-C", directory.toAbsolutePath().toString()));
        command.addAll(List.of(arguments));
        var builder = new ProcessBuilder(command).redirectErrorStream(true);
        builder.environment().put("GIT_TERMINAL_PROMPT", "0");
        builder.environment().put("GCM_INTERACTIVE", "Never");
        builder.environment().put("LC_ALL", "C");
        builder.environment().putIfAbsent("GIT_SSH_COMMAND", "ssh -o BatchMode=yes -o ConnectTimeout=15");
        Process process;
        try { process = builder.start(); }
        catch (IOException e) { throw new IOException("Git을 실행할 수 없습니다. Git 설치 및 PATH 설정을 확인하세요.", e); }
        var executor = Executors.newVirtualThreadPerTaskExecutor();
        try {
            var output = executor.submit(() -> {
                try (var input = process.getInputStream(); var kept = new ByteArrayOutputStream()) {
                    byte[] buffer = new byte[8192]; int n; boolean truncated = false;
                    while ((n = input.read(buffer)) != -1) {
                        int retain = Math.min(n, Math.max(0, 1024 * 1024 - kept.size()));
                        kept.write(buffer, 0, retain); truncated |= retain != n;
                    }
                    return kept.toString(StandardCharsets.UTF_8) + (truncated ? "\n… 출력이 길어 일부만 표시했습니다." : "");
                }
            });
            if (!process.waitFor(60, TimeUnit.SECONDS)) throw new IOException("Git 작업 시간이 초과되었습니다. 네트워크와 인증 상태를 확인하세요.");
            try { return new Result(process.exitValue(), output.get(5, TimeUnit.SECONDS)); }
            catch (ExecutionException | TimeoutException e) { throw new IOException("Git 결과를 읽지 못했습니다.", e); }
        } finally {
            process.descendants().forEach(child -> { if (child.isAlive()) child.destroyForcibly(); });
            if (process.isAlive()) process.destroyForcibly();
            try { process.getInputStream().close(); } catch (IOException ignored) { }
            executor.shutdownNow();
        }
    }
    private static Path directory(Path file) throws IOException {
        if (!Files.isRegularFile(file)) throw new IOException("문서를 먼저 저장하세요.");
        return file.toAbsolutePath().normalize().getParent();
    }
    private static Optional<Path> root(Path dir) throws IOException, InterruptedException {
        var result = run(dir, "rev-parse", "--show-toplevel");
        if (result.code() == 0) return Optional.of(Path.of(result.output().strip()).toAbsolutePath().normalize());
        if (result.output().contains("not a git repository")) return Optional.empty();
        throw new IOException(explain(result.output()));
    }
    private static Path ensureRepository(Path file) throws IOException, InterruptedException {
        var dir = directory(file); var root = root(dir);
        if (root.isPresent()) return root.get();
        require(run(dir, "init", "-b", "main")); return dir;
    }
    private static boolean hasHead(Path dir) throws IOException, InterruptedException { return run(dir, "rev-parse", "--verify", "HEAD").code() == 0; }
    public static synchronized String inspect(Path file) throws IOException, InterruptedException {
        var dir = directory(file); var repo = root(dir); var name = file.getFileName().toString();
        if (repo.isEmpty()) return "이 문서는 아직 Git으로 관리하지 않습니다.\nGit → 서버 연결 / 동기화에서 설정하거나 첫 커밋을 만드세요.\n\n문서: " + file.toAbsolutePath();
        var status = require(run(dir, "status", "--short", "--", name));
        boolean committed = hasHead(dir);
        boolean tracked = !require(run(dir, "ls-files", "--", name)).isBlank();
        String diff;
        if (!committed || !tracked) {
            var preview = new StringBuilder("새 문서 — 아직 커밋에 포함되지 않았습니다.\n");
            try (var reader = Files.newBufferedReader(file, StandardCharsets.UTF_8)) {
                String line; int count = 0;
                while ((line = reader.readLine()) != null && count++ < 1500 && preview.length() < 262144) preview.append("+ ").append(line).append('\n');
                if (line != null) preview.append("… 미리보기 일부 생략\n");
            }
            diff = preview.toString();
        } else diff = require(run(dir, "diff", "--no-ext-diff", "--no-textconv", "HEAD", "--", name));
        var log = committed ? require(run(dir, "log", "-8", "--format=%h %s", "--", name)) : "첫 커밋 전입니다.";
        return "저장소: " + repo.get() + "\n문서: " + name + "\n\n상태\n" + (status.isBlank() ? "변경 없음" : status)
                + "\n변경 내용\n" + (diff.isBlank() ? "마지막 커밋과 동일합니다." : diff) + "\n최근 기록\n" + log;
    }
    public static synchronized String commit(Path file, String message) throws IOException, InterruptedException {
        if (message.isBlank()) throw new IllegalArgumentException("커밋 메시지를 입력하세요.");
        var dir = directory(file); ensureRepository(file); var name = file.getFileName().toString();
        boolean tracked = !require(run(dir, "ls-files", "--", name)).isBlank();
        if (hasHead(dir) && tracked) {
            var diff = run(dir, "diff", "--no-ext-diff", "--no-textconv", "--quiet", "HEAD", "--", name);
            if (diff.code() == 0) return "변경 사항이 없어 새 커밋을 만들지 않았습니다.";
            if (diff.code() != 1) throw new IOException(explain(diff.output()));
        }
        if (config(dir, "user.name").isBlank() || config(dir, "user.email").isBlank())
            throw new IOException("커밋 작성자 이름과 이메일이 필요합니다. Git → 서버 연결 / 동기화에서 입력하세요.");
        require(run(dir, "add", "--", name));
        return require(run(dir, "commit", "--only", "-m", message, "--", name));
    }
    public static synchronized Settings settings(Path file) throws IOException, InterruptedException {
        var dir = directory(file); var repo = root(dir);
        if (repo.isEmpty()) return new Settings(dir, Provider.GITHUB, "", "", "", "main");
        Provider provider;
        String configuredProvider = config(dir, "gitsheet.provider");
        String origin = config(dir, "remote.origin.url");
        try { provider = Provider.valueOf(configuredProvider); }
        catch (IllegalArgumentException e) { provider = Provider.GITHUB; }
        if (configuredProvider.isBlank() && !origin.isBlank()) {
            String host = origin.toLowerCase(Locale.ROOT);
            provider = host.contains("github") ? Provider.GITHUB : host.contains("gitlab") ? Provider.GITLAB : host.contains("gitea") ? Provider.GITEA : Provider.OTHER;
        }
        return new Settings(repo.get(), provider, configuredProvider.isBlank() ? origin : config(dir, "remote." + provider.remoteName() + ".url"),
                config(dir, "user.name"), config(dir, "user.email"), branch(dir));
    }
    public static synchronized String providerUrl(Path file, Provider provider) throws IOException, InterruptedException {
        var dir = directory(file); return root(dir).isEmpty() ? "" : config(dir, "remote." + provider.remoteName() + ".url");
    }
    public static String validateUrl(String input) {
        var value = input.strip();
        if (value.isEmpty()) throw new IllegalArgumentException("원격 저장소의 Clone URL을 입력하세요.");
        if (value.chars().anyMatch(Character::isWhitespace) || value.startsWith("-")) throw new IllegalArgumentException("저장소 주소에 공백을 넣을 수 없습니다.");
        if (value.matches("[A-Za-z0-9._-]+@[A-Za-z0-9.-]+:[A-Za-z0-9_./~-]+")) return value;
        try {
            var uri = URI.create(value);
            if (!Set.of("https", "http", "ssh").contains(Objects.toString(uri.getScheme(), "")) || uri.getHost() == null || uri.getPath() == null || uri.getPath().equals("/") || uri.getPath().isEmpty()
                    || uri.getQuery() != null || uri.getFragment() != null) throw new IllegalArgumentException();
            if (uri.getUserInfo() != null && (!uri.getScheme().equals("ssh") || uri.getUserInfo().contains(":"))) throw new IllegalArgumentException();
        } catch (IllegalArgumentException e) { throw new IllegalArgumentException("HTTPS 또는 SSH Clone URL을 입력하세요. URL에 토큰·비밀번호·쿼리를 넣을 수 없습니다."); }
        return value;
    }
    public static synchronized String configure(Path file, Provider provider, String url, String name, String email) throws IOException, InterruptedException {
        url = provider == Provider.LOCAL ? "" : validateUrl(url);
        if (name.contains("\n") || email.contains("\n") || name.contains("\r") || email.contains("\r")) throw new IllegalArgumentException("작성자 정보에 줄바꿈을 넣을 수 없습니다.");
        var repo = ensureRepository(file);
        String remote = provider.remoteName();
        if (provider != Provider.LOCAL) {
            if (config(repo, "remote." + remote + ".url").isEmpty()) require(run(repo, "remote", "add", remote, url));
            else require(run(repo, "config", "--local", "--replace-all", "remote." + remote + ".url", url));
            if (!config(repo, "remote." + remote + ".pushurl").isEmpty()) require(run(repo, "config", "--local", "--unset-all", "remote." + remote + ".pushurl"));
        }
        require(run(repo, "config", "--local", "gitsheet.provider", provider.name()));
        if (!name.isBlank()) require(run(repo, "config", "--local", "user.name", name.strip()));
        if (!email.isBlank()) require(run(repo, "config", "--local", "user.email", email.strip()));
        return provider + " 연결 설정을 저장했습니다.\n" + url + "\n아직 서버로 업로드하지 않았습니다.";
    }
    private static String branch(Path dir) throws IOException, InterruptedException {
        var result = run(dir, "symbolic-ref", "--quiet", "--short", "HEAD");
        return result.code() == 0 ? result.output().strip() : "(브랜치 없음)";
    }
    private static Settings configured(Path file) throws IOException, InterruptedException {
        var s = settings(file);
        if (s.url().isBlank() || config(s.root(), "gitsheet.provider").isBlank()) throw new IOException("먼저 Git 서버 연결 설정을 저장하세요.");
        return s;
    }
    public static synchronized String checkConnection(Path file) throws IOException, InterruptedException {
        var s = configured(file); require(run(s.root(), "ls-remote", s.provider().remoteName()));
        return s.provider() + " 연결을 확인했습니다. (조회 권한 확인, 쓰기 권한은 업로드 시 확인)";
    }
    public static synchronized String fetch(Path file) throws IOException, InterruptedException {
        var s = configured(file); var result = require(run(s.root(), "fetch", s.provider().remoteName()));
        return "원격 기록을 가져왔습니다. 열린 문서와 로컬 파일은 변경하지 않았습니다.\n" + result;
    }
    public static synchronized String push(Path file) throws IOException, InterruptedException {
        var s = configured(file);
        if (!hasHead(s.root())) throw new IOException("업로드할 커밋이 없습니다. 문서를 먼저 커밋하세요.");
        if (s.branch().equals("(브랜치 없음)")) throw new IOException("브랜치를 체크아웃한 뒤 업로드하세요.");
        return require(run(s.root(), "push", "--set-upstream", s.provider().remoteName(), "HEAD:refs/heads/" + s.branch()));
    }
    private static String config(Path dir, String key) throws IOException, InterruptedException {
        var result = run(dir, "config", "--get", key);
        if (result.code() == 1) return "";
        return require(result).strip();
    }
    private static String require(Result result) throws IOException {
        if (result.code() != 0) throw new IOException(explain(result.output())); return result.output();
    }
    private static String explain(String output) {
        if (output.contains("non-fast-forward") || output.contains("fetch first")) return "서버에 다른 변경이 있습니다. 먼저 원격 기록을 가져오고 IntelliJ/Git에서 병합한 뒤 다시 업로드하세요. 강제 덮어쓰기는 하지 않았습니다.\n" + output;
        if (output.contains("Authentication failed") || output.contains("could not read Username") || output.contains("Permission denied") || output.contains("terminal prompts disabled") || output.contains("interactivity has been disabled"))
            return "Git 인증이 필요하거나 만료되었습니다. 해당 서버를 Git 자격 증명 관리자 또는 SSH 키로 로그인한 뒤 다시 시도하세요.\n" + output;
        if (output.contains("ignored")) return "이 문서는 .gitignore 규칙에 의해 제외되어 있습니다. 저장 위치 또는 제외 규칙을 확인하세요.\n" + output;
        return output.isBlank() ? "Git 명령이 실패했습니다." : output;
    }
}
