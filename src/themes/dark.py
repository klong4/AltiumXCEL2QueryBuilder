from PyQt5.QtGui import QPalette, QColor # Import QColor

class DarkTheme:
    def __init__(self):
        self.background_color = "#2E2E2E"
        self.text_color = "#FFFFFF"
        self.button_color = "#4A4A4A"
        self.button_hover_color = "#5A5A5A"
        self.border_color = "#3C3C3C"
        self.highlight_color = "#007ACC"
        self.font_family = "Arial"
        self.font_size = 10

    def apply(self, app):
        app.setStyle("Fusion")
        palette = QPalette() # Create a new palette
        # Use QColor to set colors
        palette.setColor(QPalette.Window, QColor(self.background_color))
        palette.setColor(QPalette.WindowText, QColor(self.text_color))
        palette.setColor(QPalette.Base, QColor("#1E1E1E")) # Example: Set base color slightly darker
        palette.setColor(QPalette.AlternateBase, QColor(self.background_color))
        palette.setColor(QPalette.ToolTipBase, QColor(self.background_color))
        palette.setColor(QPalette.ToolTipText, QColor(self.text_color))
        palette.setColor(QPalette.Text, QColor(self.text_color))
        palette.setColor(QPalette.Button, QColor(self.button_color))
        palette.setColor(QPalette.ButtonText, QColor(self.text_color))
        palette.setColor(QPalette.BrightText, QColor("#FF0000")) # Example: Red for bright text
        palette.setColor(QPalette.Link, QColor("#4D9EFF")) # Example: Light blue for links
        palette.setColor(QPalette.Highlight, QColor(self.highlight_color))
        palette.setColor(QPalette.HighlightedText, QColor(self.text_color))

        # Handle disabled state colors
        palette.setColor(QPalette.Disabled, QPalette.Text, QColor("#808080"))
        palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#808080"))
        palette.setColor(QPalette.Disabled, QPalette.WindowText, QColor("#808080"))
        palette.setColor(QPalette.Disabled, QPalette.Highlight, QColor("#5A5A5A"))
        palette.setColor(QPalette.Disabled, QPalette.HighlightedText, QColor("#808080"))

        app.setPalette(palette)
        # Apply minimal stylesheet for things not easily covered by palette (like borders)
        app.setStyleSheet(f"""
            QGroupBox {{ 
                border: 1px solid {self.border_color}; 
                margin-top: 0.5em; 
            }}
            QGroupBox::title {{ 
                subcontrol-origin: margin; 
                left: 10px; 
                padding: 0 3px 0 3px; 
            }}
            QTabWidget::pane {{
                border: 1px solid {self.border_color};
            }}
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
                border: 1px solid {self.border_color};
                padding: 2px;
            }}
            QTableView {{
                 gridline-color: {self.border_color};
            }}
            /* Style buttons slightly differently if needed, as palette handles base color */
            QPushButton {{
                border: 1px solid {self.border_color};
                padding: 3px;
            }}
            QPushButton:hover {{
                background-color: {self.button_hover_color};
            }}
            QPushButton:disabled {{
                border: 1px solid #5A5A5A; /* Match disabled highlight */
            }}
        """)