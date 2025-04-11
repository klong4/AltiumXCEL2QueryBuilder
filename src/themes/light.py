class LightTheme:
    def __init__(self):
        self.background_color = "#FFFFFF"
        self.text_color = "#000000"
        self.button_color = "#E0E0E0"
        self.button_hover_color = "#D0D0D0"
        self.border_color = "#CCCCCC"
        self.accent_color = "#007BFF"

    def apply(self, app):
        app.setStyleSheet(f"""
            QWidget {{
                background-color: {self.background_color};
                color: {self.text_color};
            }}
            QPushButton {{
                background-color: {self.button_color};
                border: 1px solid {self.border_color};
            }}
            QPushButton:hover {{
                background-color: {self.button_hover_color};
            }}
            QTabWidget::pane {{
                border: 1px solid {self.border_color};
            }}
            QLineEdit {{
                border: 1px solid {self.border_color};
                padding: 5px;
            }}
            QLabel {{
                color: {self.text_color};
            }}
        """)