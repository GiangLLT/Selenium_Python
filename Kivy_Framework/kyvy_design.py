# pip install kivy
# pip install buildozer => deploy apk
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout

class MyApp(App):
    def build(self):
        layout = BoxLayout(orientation='horizontal')
        self.label = Label(text="Hello, Kivy!")
        self.button = Button(text="Click Me")
        self.button.bind(on_press=self.on_button_press)

        layout.add_widget(self.label)
        layout.add_widget(self.button)

        return layout

    def on_button_press(self, instance):
        self.label.text = "Button Clicked"
        self.button.text = "Change text button"

if __name__ == "__main__":
    MyApp().run()
