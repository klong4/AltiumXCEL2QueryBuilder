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
        palette = app.palette()
        palette.setColor(palette.Window, self.background_color)
        palette.setColor(palette.WindowText, self.text_color)
        palette.setColor(palette.Button, self.button_color)
        palette.setColor(palette.ButtonText, self.text_color)
        palette.setColor(palette.Highlight, self.highlight_color)
        app.setPalette(palette)