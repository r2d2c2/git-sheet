package io.github.r2d2c2.gitsheet;

import javafx.animation.PauseTransition;
import javafx.geometry.*;
import javafx.scene.control.Label;
import javafx.scene.layout.*;
import javafx.util.Duration;
import java.nio.file.Path;

/** In-window popup: never steals focus, repeated saves restart its four-second timer. */
final class SaveNotification extends VBox {
    static final Duration DISPLAY_TIME = Duration.seconds(4);
    private final Label message = new Label();
    private final PauseTransition timeout = new PauseTransition(DISPLAY_TIME);
    SaveNotification(StackPane host) {
        super(5); getStyleClass().add("save-notification");
        var heading = new Label("저장 완료"); heading.setStyle("-fx-font-weight:bold; -fx-text-fill:white;");
        message.setWrapText(true); message.setStyle("-fx-text-fill:white;");
        getChildren().addAll(heading, message); setPadding(new Insets(14, 18, 14, 18));
        setMaxSize(620, USE_PREF_SIZE); setMouseTransparent(true); setVisible(false); setManaged(false);
        StackPane.setAlignment(this, Pos.BOTTOM_RIGHT); StackPane.setMargin(this, new Insets(0, 20, 48, 20)); host.getChildren().add(this);
        timeout.setOnFinished(_ -> dismiss());
    }
    void show(Path path) {
        message.setText(path.toAbsolutePath().normalize().toString()); setManaged(true); setVisible(true); toFront(); timeout.playFromStart();
    }
    String displayedPath() { return message.getText(); }
    void dismiss() { timeout.stop(); setVisible(false); setManaged(false); }
}
