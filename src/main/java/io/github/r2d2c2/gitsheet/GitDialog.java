package io.github.r2d2c2.gitsheet;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.Window;
import java.nio.file.Path;
import java.util.EnumMap;
import java.util.concurrent.*;

/** Explicit remote selection and operations, with no credential collection in workbook files. */
final class GitDialog extends Dialog<Void> {
    private final Path file;
    private final Executor executor;
    private final ComboBox<GitService.Provider> provider = new ComboBox<>(FXCollections.observableArrayList(GitService.Provider.values()));
    private final TextField url = new TextField(), name = new TextField(), email = new TextField();
    private final TextArea result = new TextArea();
    private final Label location = new Label("설정 불러오는 중…");
    private final VBox controls = new VBox(10);
    private final EnumMap<GitService.Provider, String> drafts = new EnumMap<>(GitService.Provider.class);
    private boolean loading, busy;
    private GitService.Settings current;
    private CompletableFuture<Void> idle = CompletableFuture.completedFuture(null);
    CompletableFuture<Void> whenIdle() { return idle; }
    GitDialog(Window owner, Path file, Executor executor) {
        this.file = file; this.executor = executor; initOwner(owner); setTitle("Git 서버 연결 / 동기화"); setResizable(true);
        getDialogPane().getButtonTypes().add(ButtonType.CLOSE);
        location.setWrapText(true);
        var form = new GridPane(); form.setHgap(12); form.setVgap(10);
        provider.setMaxWidth(Double.MAX_VALUE); url.setPrefColumnCount(44);
        provider.setId("git-provider"); url.setId("git-url"); name.setId("git-name"); email.setId("git-email");
        form.addRow(0, new Label("서비스"), provider);
        form.addRow(1, new Label("Clone URL"), url);
        form.addRow(2, new Label("커밋 이름"), name);
        form.addRow(3, new Label("커밋 이메일"), email);
        GridPane.setHgrow(url, Priority.ALWAYS);
        var hint = new Label("GitHub · GitLab · Gitea 및 자체 서버의 HTTPS/SSH 주소를 지원합니다.\n로그인은 Git 자격 증명 관리자 또는 SSH 키를 사용합니다. 이름·이메일은 이 저장소에만 저장합니다."); hint.setWrapText(true);
        var save = button("설정 저장", () -> execute("설정 저장", () -> saveSettings()));
        save.setId("git-save");
        var check = button("연결 확인", () -> network("연결 확인", () -> GitService.checkConnection(file)));
        var push = button("업로드 (push)", this::push);
        var fetch = button("원격 기록 가져오기", () -> network("원격 기록 가져오기", () -> GitService.fetch(file)));
        result.setEditable(false); result.setWrapText(true); result.setPrefRowCount(10);
        controls.getChildren().addAll(location, form, hint, new HBox(8, save, check, push, fetch));
        var content = new VBox(12, controls, result); content.setPadding(new Insets(10)); content.setPrefWidth(720);
        getDialogPane().setContent(content);
        provider.valueProperty().addListener((_, old, next) -> {
            if (loading || next == null) return;
            if (old != null) drafts.put(old, url.getText());
            url.setText(drafts.getOrDefault(next, "")); url.setPromptText(next.example());
            url.setDisable(next == GitService.Provider.LOCAL);
            check.setDisable(next == GitService.Provider.LOCAL); push.setDisable(next == GitService.Provider.LOCAL); fetch.setDisable(next == GitService.Provider.LOCAL);
        });
        setOnCloseRequest(e -> { if (busy) e.consume(); });
        execute("연결 설정 읽기", () -> {
            var settings = GitService.settings(file);
            var urls = new EnumMap<GitService.Provider, String>(GitService.Provider.class);
            for (var p : GitService.Provider.values()) urls.put(p, GitService.providerUrl(file, p));
            Platform.runLater(() -> {
                current = settings; drafts.putAll(urls); loading = true;
                provider.setValue(settings.provider()); url.setText(settings.url()); url.setPromptText(settings.provider().example());
                name.setText(settings.name()); email.setText(settings.email()); loading = false;
                boolean local = settings.provider() == GitService.Provider.LOCAL;
                url.setDisable(local); check.setDisable(local); push.setDisable(local); fetch.setDisable(local);
                location.setText("문서: " + file.toAbsolutePath() + "\n저장소: " + settings.root() + "\n브랜치: " + settings.branch());
            });
            return "서비스와 Clone URL을 선택한 뒤 설정을 저장하세요. 기존 origin은 유지됩니다.\n서버 저장소는 미리 만들어 두어야 합니다.";
        });
    }
    private static Button button(String text, Runnable action) { var b = new Button(text); b.setOnAction(_ -> action.run()); return b; }
    private record Form(GitService.Provider provider, String url, String name, String email) {}
    private Form form() { return new Form(provider.getValue(), url.getText(), name.getText(), email.getText()); }
    // Capture controls on the FX thread; background tasks never read UI state.
    private Form pending;
    private String saveSettings() throws Exception { return GitService.configure(file, pending.provider(), pending.url(), pending.name(), pending.email()); }
    private void network(String title, Operation operation) { execute(title, () -> { saveSettings(); return operation.run(); }); }
    private void push() {
        var target = form();
        try { GitService.validateUrl(target.url()); } catch (IllegalArgumentException e) { result.setText(e.getMessage()); return; }
        var confirm = new Alert(Alert.AlertType.CONFIRMATION,
                "대상: " + target.provider() + "\n" + target.url() + "\n\n저장소: " + (current == null ? file.getParent() : current.root())
                        + "\n현재 브랜치의 모든 커밋을 업로드합니다. 열린 문서의 미커밋 변경은 포함되지 않습니다.", ButtonType.OK, ButtonType.CANCEL);
        confirm.initOwner(getDialogPane().getScene().getWindow()); confirm.setHeaderText("이 서버에 업로드할까요?");
        if (confirm.showAndWait().orElse(ButtonType.CANCEL) == ButtonType.OK) network("업로드", () -> GitService.push(file));
    }
    @FunctionalInterface private interface Operation { String run() throws Exception; }
    private void execute(String title, Operation operation) {
        if (busy) return;
        idle = new CompletableFuture<>();
        pending = form(); busy = true; controls.setDisable(true); getDialogPane().lookupButton(ButtonType.CLOSE).setDisable(true); result.setText(title + " 중…");
        CompletableFuture.supplyAsync(() -> { try { return operation.run(); } catch (Exception e) { throw new CompletionException(e); } }, executor)
                .whenComplete((text, failure) -> Platform.runLater(() -> {
                    busy = false; controls.setDisable(false); getDialogPane().lookupButton(ButtonType.CLOSE).setDisable(false);
                    result.setText(failure == null ? text : "실패: " + failure.getCause().getMessage());
                    if (failure == null) idle.complete(null); else idle.completeExceptionally(failure);
                }));
    }
}
