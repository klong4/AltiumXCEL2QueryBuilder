from PyQt5.QtGui import QPalette, QColor # Import QColor

class LightTheme:
    def __init__(self):
        self.window_color = "#F0F0F0" # General window background
        self.base_color = "#FFFFFF"   # Text entry background
        self.text_color = "#000000"
        self.button_color = "#E0E0E0"
        self.border_color = "#B0B0B0" # Slightly darker border
        self.highlight_color = "#007BFF" # Accent/Highlight
        self.highlighted_text_color = "#FFFFFF"
        self.disabled_text_color = "#A0A0A0"
        self.disabled_button_text_color = "#707070"
        self.disabled_highlight_color = "#D0D0D0"

    def apply(self, app):
        app.setStyle("Fusion") # Use Fusion style for consistency
        palette = QPalette()

        # Define colors using QColor
        palette.setColor(QPalette.Window, QColor(self.window_color))
        palette.setColor(QPalette.WindowText, QColor(self.text_color))
        palette.setColor(QPalette.Base, QColor(self.base_color))
        palette.setColor(QPalette.AlternateBase, QColor("#E8E8E8")) # Slightly off-white for alternates
        palette.setColor(QPalette.ToolTipBase, QColor(self.base_color))
        palette.setColor(QPalette.ToolTipText, QColor(self.text_color))
        palette.setColor(QPalette.Text, QColor(self.text_color))
        palette.setColor(QPalette.Button, QColor(self.button_color))
        palette.setColor(QPalette.ButtonText, QColor(self.text_color))
        palette.setColor(QPalette.BrightText, QColor("#FF0000")) # Red for bright text (standard)
        palette.setColor(QPalette.Link, QColor(self.highlight_color))
        palette.setColor(QPalette.Highlight, QColor(self.highlight_color))
        palette.setColor(QPalette.HighlightedText, QColor(self.highlighted_text_color))

        # Handle disabled state colors
        palette.setColor(QPalette.Disabled, QPalette.Text, QColor(self.disabled_text_color))
        palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(self.disabled_button_text_color))
        palette.setColor(QPalette.Disabled, QPalette.WindowText, QColor(self.disabled_text_color))
        palette.setColor(QPalette.Disabled, QPalette.Highlight, QColor(self.disabled_highlight_color))
        palette.setColor(QPalette.Disabled, QPalette.HighlightedText, QColor(self.text_color)) # Keep text readable on disabled highlight

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
        """)