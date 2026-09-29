package io.github.r2d2c2.gitsheet;

/** Separate launcher keeps JavaFX module discovery compatible with IDE execution. */
public final class Launcher {
    private Launcher() {}
    public static void main(String[] args) { SpreadsheetApp.main(args); }
}
