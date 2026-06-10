module org.example {
    requires java.desktop;
    requires javafx.controls;
    requires javafx.fxml;
    requires org.controlsfx.controls;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;

    opens org.example to javafx.fxml;
    exports org.example;
}
