# key_presser.py

import time
import ctypes
import pynput

class KeyPresser:
    """
    A class to handle key press and release actions on Windows using Direct Input.
    """
    _instance = None
    # Key Codes found at: https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-6.0/aa299374(v=vs.60)
    KEY_CODES = {
        "Q": 0x10,
        "E": 0x12,
        "W": 0x11,
        "R": 0x13,
        "T": 0x14,
        "Y": 0x15,
        "U": 0x16,
        "I": 0x17,
        "O": 0x18,
        "P": 0x19,
        "A": 0x1E,
        "S": 0x1F,
        "D": 0x20,
        "F": 0x21,
        "G": 0x22,
        "H": 0x23,
        "J": 0x24,
        "K": 0x25,
        "L": 0x26,
        "Z": 0x2C,
        "X": 0x2D,
        "C": 0x2E,
        "V": 0x2F,
        "B": 0x30,
        "N": 0x31,
        "M": 0x32,
        "LEFT_ARROW": 0xCB,
        "RIGHT_ARROW": 0xCD,
        "UP_ARROW": 0xC8,
        "DOWN_ARROW": 0xD0,
        "ESC": 0x01,
        "ONE": 0x02,
        "TWO": 0x03,
        "THREE": 0x04,
        "FOUR": 0x05,
        "FIVE": 0x06,
        "SIX": 0x07,
        "SEVEN": 0x08,
        "EIGHT": 0x09,
        "NINE": 0x0A,
        "ZERO": 0x0B,
        "MINUS": 0x0C,
        "EQUALS": 0x0D,
        "BACKSPACE": 0x0E,
        "APOSTROPHE": 0x28,
        "SEMICOLON": 0x27,
        "TAB": 0x0F,
        "CAPSLOCK": 0x3A,
        "ENTER": 0x1C,
        "LEFT_CONTROL": 0x1D,
        "LEFT_ALT": 0x38,
        "LEFT_SHIFT": 0x2A,
        "RIGHT_SHIFT": 0x36,
        "TILDE": 0x29,
        "PRINTSCREEN": 0x37,
        "NUM_LOCK": 0x45,
        "SPACE": 0x39,
        "DELETE": 0x53,
        "COMMA": 0x33,
        "PERIOD": 0x34,
        "BACKSLASH": 0x35,
        "FORWARDSLASH": 0x2B,
        "LEFT_BRACKET": 0x1A,
        "RIGHT_BRACKET": 0x1B,
        "F1": 0x3B,
        "F2": 0x3C,
        "F3": 0x3D,
        "F4": 0x3E,
        "F5": 0x3F,
        "F6": 0x40,
        "F7": 0x41,
        "F8": 0x42,
        "F9": 0x43,
        "F10": 0x44,
        "F11": 0x57,
        "F12": 0x58,
        "NUMPAD_0": 0x52,
        "NUMPAD_1": 0x4F,
        "NUMPAD_2": 0x50,
        "NUMPAD_3": 0x51,
        "NUMPAD_4": 0x4B,
        "NUMPAD_5": 0x4C,
        "NUMPAD_6": 0x4D,
        "NUMPAD_7": 0x47,
        "NUMPAD_8": 0x48,
        "NUMPAD_9": 0x49,
        "NUMPAD_PLUS": 0x4E,
        "NUMPAD_MINUS": 0x4A,
        "NUMPAD_PERIOD": 0x53,
        "NUMPAD_ENTER": 0x9C,
        "NUMPAD_BACKSLASH": 0xB5,
        "LEFT_MOUSE": 0x100,
        "RIGHT_MOUSE": 0x101,
        "MIDDLE_MOUSE": 0x102,
        "MOUSE3": 0x103,
        "MOUSE4": 0x104,
        "MOUSE5": 0x105,
        "MOUSE6": 0x106,
        "MOUSE7": 0x107,
        "MOUSE_WHEEL_UP": 0x108,
        "MOUSE_WHEEL_DOWN": 0x109
    }
    KEY_EVENT_KEY_UP = 0x0002
    KEY_EVENT_SCANCODE = 0x0008
    INPUT_KEYBOARD = 1

    def __init__(self) -> None:
        self.pressed_keys = set()

        # Direct Input functions found at: https://stackoverflow.com/questions/53643273/how-to-keep-pynput-and-ctypes-from-clashing
        # Use these to prevent conflict errors with pynput.
        self.send_input = ctypes.windll.user32.SendInput

    @classmethod
    def get_instance(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def get_hex_key_code(self, key_name: str) -> int:
        """
        Returns the hex code for the given key.

        Args:
            key_name (str): The name of the key to be searched.

        Returns:
            int: the hex code for the given key.

        Raises:
            ValueError: If key_nam,e is not found in the KEY_CODES.
        """
        key_name = key_name.upper()
        if key_name not in self.KEY_CODES:
            valid_keys = ', '.join(self.KEY_CODES.keys())
            raise ValueError(
                f"{key_name} is not a valid key name. "
                "Please choose from the following valid keys: "
                f"{valid_keys}"
            )
        hex_code = self.KEY_CODES[key_name]
        return hex_code

    def hold_key(self, key_name: str) -> None:
        """
        Holds down the key with the given name.

        Args:
            key_name: The name of the key to be held down.

        Returns:
            None
        """
        hex_key_code = self.get_hex_key_code(key_name)
        extra_info = ctypes.c_ulong(0)
        input_union = pynput._util.win32.INPUT_union()
        input_union.ki = pynput._util.win32.KEYBDINPUT(0, hex_key_code, self.KEY_EVENT_SCANCODE, 0, ctypes.cast(ctypes.pointer(extra_info), ctypes.c_void_p))
        keyboard_input = pynput._util.win32.INPUT(ctypes.c_ulong(self.INPUT_KEYBOARD), input_union)
        self.send_input(1, ctypes.pointer(keyboard_input), ctypes.sizeof(keyboard_input))
        self.pressed_keys.add(key_name)

    def release_key(self, key_name: str) -> None:
        """
        Releases the key with the given name.

        Args:
            key_name: The name of the key to be released.

        Returns:
            None
        """
        hex_key_code = self.get_hex_key_code(key_name)
        extra_info = ctypes.c_ulong(0)
        input_union = pynput._util.win32.INPUT_union()
        input_union.ki = pynput._util.win32.KEYBDINPUT(0, hex_key_code, self.KEY_EVENT_SCANCODE | self.KEY_EVENT_KEY_UP, 0, ctypes.cast(ctypes.pointer(extra_info), ctypes.c_void_p))
        keyboard_input = pynput._util.win32.INPUT(ctypes.c_ulong(self.INPUT_KEYBOARD), input_union)
        self.send_input(1, ctypes.pointer(keyboard_input), ctypes.sizeof(keyboard_input))
        self.pressed_keys.discard(key_name)

    def hold_and_release_key(self, key_name: str, seconds: float) -> None:
        """
        Holds down the key with the given name for a specified duration and then releases it.

        Args:
            key_name: The name of the key to be held down and released.
            seconds: The duration in seconds for which the key should be held down.

        Returns:
            None
        """
        self.hold_key(key_name)
        time.sleep(seconds)
        self.release_key(key_name)

    def release_all_keys(self) -> None:
        """
        Releases all currently pressed keys.

        Args:
            None

        Returns:
            None
        """
        for key_code in list(self.pressed_keys):
            self.release_key(key_code)
