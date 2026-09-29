package io.github.r2d2c2.gitsheet;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;

/** Executes argument arrays, never a shell. Only the selected document is staged/committed. */
public final class GitService {
    private GitService() {}
    public record Result(int code, String output) {}
    public static Result run(Path directory, String... arguments) throws IOException, InterruptedException {
        var command = new ArrayList<>(List.of("git", "-C", directory.toAbsolutePath().toString()));
        command.addAll(List.of(arguments));
        var builder = new ProcessBuilder(command).redirectErrorStream(true);
        builder.environment().put("GIT_TERMINAL_PROMPT", "0");
        var process = builder.start();
        try (var executor = Executors.newVirtualThreadPerTaskExecutor()) {
            var output = executor.submit(() -> new String(process.getInputStream().readAllBytes(), StandardCharsets.UTF_8));
            if (!process.waitFor(30, TimeUnit.SECONDS)) { process.destroyForcibly(); throw new IOException("Git 작업 시간이 초과되었습니다."); }
            try { return new Result(process.exitValue(), output.get(5, TimeUnit.SECONDS)); }
            catch (ExecutionException | TimeoutException e) { throw new IOException(e); }
        } finally { if (process.isAlive()) process.destroyForcibly(); }
    }
    public static String inspect(Path file) throws IOException, InterruptedException {
        var dir = file.toAbsolutePath().getParent(); var name = file.getFileName().toString();
        var status = run(dir, "status", "--short", "--", name);
        if (status.code() != 0) return status.output();
        var diff = run(dir, "diff", "HEAD", "--", name);
        var log = run(dir, "log", "-8", "--format=%h %s", "--", name);
        return "상태\n" + status.output() + "\n변경 내용\n" + diff.output() + "\n최근 기록\n" + log.output();
    }
    public static String commit(Path file, String message) throws IOException, InterruptedException {
        if (message.isBlank()) throw new IllegalArgumentException("커밋 메시지를 입력하세요.");
        var dir = file.toAbsolutePath().getParent(); var name = file.getFileName().toString();
        if (run(dir, "rev-parse", "--show-toplevel").code() != 0) require(run(dir, "init", "-b", "main"));
        require(run(dir, "add", "--", name));
        return require(run(dir, "commit", "--only", "-m", message, "--", name));
    }
    private static String require(Result result) throws IOException {
        if (result.code() != 0) throw new IOException(result.output()); return result.output();
    }
}
