import time
import keyboard
from pynput import mouse
import os

def record_user_actions(output_file):
    print("Recording user actions. Press 'q' to stop recording.")
    actions = []
    
    # Define keyboard event to stop recording
    keyboard.add_hotkey('q', lambda: actions.append("EndRecording"))

    # Define mouse event listener
    def on_click(x, y, button, pressed):
        if pressed:
            actions.append(f"Click at ({x}, {y})")

    mouse_listener = mouse.Listener(on_click=on_click)
    mouse_listener.start()

    # Record actions until 'q' is pressed
    while True:
        if "EndRecording" in actions:
            break
        time.sleep(0.1)  # Adjust sleep time based on recording frequency

    mouse_listener.stop()
    keyboard.remove_hotkey('q')

    # Create directory if it doesn't exist
    directory = os.path.dirname(output_file)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Save recorded actions to file
    with open(output_file, 'w') as f:
        for action in actions:
            f.write(action + '\n')

    print(f"Recorded user actions saved to {output_file}")

def main():
    output_file = '.\\Application\\Script\\Script.txt'
    record_user_actions(output_file)

if __name__ == "__main__":
    main()
