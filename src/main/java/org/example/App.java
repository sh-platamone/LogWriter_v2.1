package org.example;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class App extends Application {

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader loader = new FXMLLoader(App.class.getResource("schedule-workbench.fxml"));
        Scene scene = new Scene(loader.load(), 1400, 940);
        scene.getStylesheets().add(App.class.getResource("app.css").toExternalForm());
        stage.setScene(scene);
        stage.setTitle("LogWriter");
        stage.setMinWidth(1100);
        stage.setMinHeight(700);
        stage.show();
    }

    public static void main(String[] args) {
        launch();
    }
}
