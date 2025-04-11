class ThemeManager:
    def __init__(self, app):
        self.app = app
        self.current_theme = None

    def apply_theme(self, theme):
        if theme == 'dark':
            self.load_dark_theme()
        elif theme == 'light':
            self.load_light_theme()

    def load_dark_theme(self):
        with open('resources/styles/dark.qss', 'r') as file:
            style = file.read()
            self.app.setStyleSheet(style)
            self.current_theme = 'dark'

    def load_light_theme(self):
        with open('resources/styles/light.qss', 'r') as file:
            style = file.read()
            self.app.setStyleSheet(style)
            self.current_theme = 'light'

    def toggle_theme(self):
        if self.current_theme == 'dark':
            self.load_light_theme()
        else:
            self.load_dark_theme()