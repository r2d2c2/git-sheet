package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.*;
import static org.junit.jupiter.api.Assertions.*;

class GitServiceTest {
    @TempDir Path directory;
    private void init() throws Exception {
        assertEquals(0, GitService.run(directory, "init", "-b", "main").code());
        GitService.run(directory, "config", "user.name", "Git Sheet Test");
        GitService.run(directory, "config", "user.email", "test@example.invalid");
        GitService.run(directory, "config", "commit.gpgsign", "false");
    }
    @Test void commitsOnlySelectedFileAndLeavesOtherStagedChangesAlone() throws Exception {
        init(); var document = directory.resolve("한글 book.gsheet"); Files.writeString(document, "first\n");
        Files.writeString(directory.resolve("unrelated.txt"), "private\n"); GitService.run(directory, "add", "unrelated.txt");
        GitService.commit(document, "Save workbook");
        var committed = GitService.run(directory, "show", "--format=", "--name-only", "HEAD").output();
        assertFalse(committed.contains("unrelated.txt"));
        assertTrue(GitService.run(directory, "diff", "--cached", "--name-only").output().contains("unrelated.txt"));
        Files.writeString(document, "second\n"); var diff = GitService.inspect(document);
        assertTrue(diff.contains("-first")); assertTrue(diff.contains("+second"));
        assertTrue(diff.contains("Save workbook"));
    }
    @Test void passesMetacharactersAsLiteralCommitMessage() throws Exception {
        init(); var document = directory.resolve("book.gsheet"); Files.writeString(document, "x");
        GitService.commit(document, "Literal $(echo hello); & <tag>");
        assertEquals("Literal $(echo hello); & <tag>", GitService.run(directory, "log", "-1", "--format=%s").output().trim());
    }
    @Test void handlesNonRepositoryAndUnbornHeadWithoutFatalErrors() throws Exception {
        var file = directory.resolve("new.gsheet"); Files.writeString(file, "one\ntwo\n");
        assertTrue(GitService.inspect(file).contains("아직 Git"));
        init(); var preview = GitService.inspect(file);
        assertTrue(preview.contains("첫 커밋 전")); assertTrue(preview.contains("+ one")); assertFalse(preview.contains("fatal:"));
    }
    @Test void unchangedCommitIsSuccessfulNoOp() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "same");
        GitService.commit(file, "first"); var before = GitService.run(directory, "rev-parse", "HEAD").output();
        assertTrue(GitService.commit(file, "again").contains("변경 사항이 없어"));
        assertEquals(before, GitService.run(directory, "rev-parse", "HEAD").output());
    }
    @Test void treatsBracketFilenamesLiterally() throws Exception {
        init(); var selected = directory.resolve("[a].gsheet"); Files.writeString(selected, "chosen");
        Files.writeString(directory.resolve("a.gsheet"), "unrelated");
        GitService.commit(selected, "literal path");
        var files = GitService.run(directory, "ls-tree", "--name-only", "HEAD").output();
        assertEquals("[a].gsheet", files.strip());
        assertTrue(GitService.inspect(directory.resolve("a.gsheet")).contains("+ unrelated"));
    }
    @Test void missingIdentityIsActionableAndDoesNotStageDocument() throws Exception {
        init(); GitService.run(directory, "config", "--local", "user.name", "");
        var file = directory.resolve("book.gsheet"); Files.writeString(file, "x");
        var failure = assertThrows(java.io.IOException.class, () -> GitService.commit(file, "test"));
        assertTrue(failure.getMessage().contains("커밋 작성자"));
        assertTrue(GitService.run(directory, "ls-files").output().isBlank());
    }
    @Test void providerSettingsPersistAndNeverReplaceOrigin() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "x");
        GitService.run(directory, "remote", "add", "origin", "https://example.com/existing.git");
        for (var provider : GitService.Provider.values()) {
            var url = provider == GitService.Provider.LOCAL ? "" : "https://example.com/owner/" + provider.name() + ".git";
            GitService.configure(file, provider, url, "Local Author", "local@example.invalid");
            var settings = GitService.settings(file); assertEquals(provider, settings.provider()); assertEquals(url, settings.url());
            assertEquals("Local Author", settings.name());
        }
        assertEquals("https://example.com/existing.git", GitService.run(directory, "remote", "get-url", "origin").output().strip());
        assertTrue(GitService.providerUrl(file, GitService.Provider.GITLAB).contains("GITLAB.git"));
    }
    @Test void savingConnectionUsesOneExplicitDestination() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "x");
        GitService.configure(file, GitService.Provider.GITHUB, "https://github.com/one/repo.git", "", "");
        GitService.run(directory, "config", "--add", "remote.gitsheet-github.url", "https://example.com/extra.git");
        GitService.run(directory, "config", "remote.gitsheet-github.pushurl", "https://example.com/wrong.git");
        GitService.configure(file, GitService.Provider.GITHUB, "https://github.com/two/repo.git", "", "");
        assertEquals("https://github.com/two/repo.git", GitService.run(directory, "remote", "get-url", "--push", "--all", "gitsheet-github").output().strip());
    }
    @Test void supportsHttpsSshAndSelfHostedUrlsButRejectsTokensAndHelpers() {
        for (String url : new String[]{"https://github.com/me/repo.git", "https://gitlab.example/group/sub/repo.git", "http://localhost:3000/me/repo.git", "ssh://git@gitea.example:2222/me/repo.git", "git@github.com:me/repo.git"}) assertEquals(url, GitService.validateUrl(url));
        for (String url : new String[]{"", "-evil", "ext::sh -c command", "file:///tmp/repo", "https://token@github.com/me/repo.git", "https://github.com/me/repo?token=secret", "ssh://git:secret@host/repo", "https://github.com/"}) assertThrows(IllegalArgumentException.class, () -> GitService.validateUrl(url));
    }
    @Test void pushesAndFetchesUsingEachSelectedProvider() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "first"); GitService.commit(file, "first");
        for (var provider : new GitService.Provider[]{GitService.Provider.GITHUB, GitService.Provider.GITLAB, GitService.Provider.GITEA}) {
            var remote = Files.createDirectory(directory.resolve(provider.name() + ".git"));
            assertEquals(0, GitService.run(remote, "init", "--bare", "-b", "main").code());
            GitService.configure(file, provider, "https://example.com/owner/repo.git", "", "");
            // Local bare remotes exercise Git's real transfer protocol without external accounts.
            GitService.run(directory, "remote", "set-url", provider.remoteName(), remote.toString());
            assertTrue(GitService.checkConnection(file).contains("연결을 확인"));
            GitService.push(file);
            assertEquals(GitService.run(directory, "rev-parse", "HEAD").output(), GitService.run(remote, "rev-parse", "main").output());
            assertTrue(GitService.fetch(file).contains("가져왔습니다"));
        }
    }
    @Test void refusesPushBeforeFirstCommit() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "x");
        GitService.configure(file, GitService.Provider.GITEA, "https://example.com/me/repo.git", "", "");
        assertTrue(assertThrows(java.io.IOException.class, () -> GitService.push(file)).getMessage().contains("커밋이 없습니다"));
    }
    @Test void prefillsExistingOriginWithoutChangingIt() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "x");
        GitService.run(directory, "remote", "add", "origin", "https://gitlab.example/team/repo.git");
        var settings = GitService.settings(file);
        assertEquals(GitService.Provider.GITLAB, settings.provider()); assertEquals("https://gitlab.example/team/repo.git", settings.url());
        assertTrue(assertThrows(java.io.IOException.class, () -> GitService.push(file)).getMessage().contains("설정을 저장"));
    }
    @Test void rejectsDivergedPushWithoutOverwritingRemoteHistory() throws Exception {
        init(); var file = directory.resolve("book.gsheet"); Files.writeString(file, "base"); GitService.commit(file, "base");
        var remote = Files.createDirectory(directory.resolve("server.git")); GitService.run(remote, "init", "--bare", "-b", "main");
        GitService.configure(file, GitService.Provider.GITLAB, "https://example.com/team/repo.git", "", "");
        GitService.run(directory, "remote", "set-url", "gitsheet-gitlab", remote.toString()); GitService.push(file);
        var peer = directory.resolve("peer"); assertEquals(0, GitService.run(directory, "clone", remote.toString(), peer.toString()).code());
        GitService.run(peer, "config", "user.name", "Peer"); GitService.run(peer, "config", "user.email", "peer@example.invalid"); GitService.run(peer, "config", "commit.gpgsign", "false");
        Files.writeString(peer.resolve("book.gsheet"), "server change"); GitService.commit(peer.resolve("book.gsheet"), "server change"); GitService.run(peer, "push", "origin", "main");
        var serverHead = GitService.run(remote, "rev-parse", "main").output();
        Files.writeString(file, "local change"); GitService.commit(file, "local change");
        assertTrue(assertThrows(java.io.IOException.class, () -> GitService.push(file)).getMessage().contains("서버에 다른 변경"));
        assertEquals(serverHead, GitService.run(remote, "rev-parse", "main").output());
    }
}
