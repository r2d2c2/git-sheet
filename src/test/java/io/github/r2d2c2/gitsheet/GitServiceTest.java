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
}
