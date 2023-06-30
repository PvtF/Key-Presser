# Key Presser

KeyPresser is a Python utility module that provides an interface for simulating key press and release actions in Windows using the Direct Input method. This functionality is extremely useful for a wide range of applications, including automating repetitive tasks, testing software, or creating unique projects such as a "Twitch Plays" style game. As a singleton class, KeyPresser ensures only a single instance of key pressing functionality exists throughout your application, preventing possible conflicts or redundant operations.

This is an adaptation of the "TwitchPlays_KeyCodes"
https://github.com/DougDougGithub/TwitchPlays

## Using Key Presser

KeyPresser is designed to be simple and straightforward to use. Here are a few examples of the functions available:

**Getting the Singleton Instance:**
To get the singleton instance of the KeyPresser class, simply call the _get_instance()_ method:

```python
import KeyPresser

kp = KeyPresser.get_instance()
```

**Holding and Releasing a Key:**
To simulate holding a key for a certain duration and then releasing it, you can use the _hold_and_release_key()_ method. 
The following example presses and holds the "W" key for 0.5 seconds:

```python
kp.hold_and_release_key('W', 0.5)
```

**Holding a Key Indefinitely:**
If you need to hold a key indefinitely, use the _hold_key()_ method. The key will remain held until you explicitly release it:

```python
kp.hold_key('W')  # This will hold the 'W' key
# ... Some other code ...
kp.release_key('W')  # This will release the 'W' key
```

**Releasing All Keys:**
If you want to ensure all keys are released, regardless of their state, you can use the _release_all_keys()_ method:

```python
kp.release_all_keys()
```

It's important to note that the key names passed to these methods should match the ones defined in the KEY_CODES dictionary in the KeyPresser class. For a list of valid key names, you can refer to this dictionary.
|Key Name|Hex Code|
|-----------------|-----------------|
| Q | 0x10 |
| E | 0x12 |
| W | 0x11 |
| R | 0x13 |
| T | 0x14 |
| Y | 0x15 |
| U | 0x16 |
| I | 0x17 |
| O | 0x18 |
| P | 0x19 |
| A | 0x1e |
| S | 0x1f |
| D | 0x20 |
| F | 0x21 |
| G | 0x22 |
| H | 0x23 |
| J | 0x24 |
| K | 0x25 |
| L | 0x26 |
| Z | 0x2c |
| X | 0x2d |
| C | 0x2e |
| V | 0x2f |
| B | 0x30 |
| N | 0x31 |
| M | 0x32 |
| LEFT_ARROW | 0xcb |
| RIGHT_ARROW | 0xcd |
| UP_ARROW | 0xc8 |
| DOWN_ARROW | 0xd0 |
| ESC | 0x1 |
| ONE | 0x2 |
| TWO | 0x3 |
| THREE | 0x4 |
| FOUR | 0x5 |
| FIVE | 0x6 |
| SIX | 0x7 |
| SEVEN | 0x8 |
| EIGHT | 0x9 |
| NINE | 0xa |
| ZERO | 0xb |
| MINUS | 0xc |
| EQUALS | 0xd |
| BACKSPACE | 0xe |
| APOSTROPHE | 0x28 |
| SEMICOLON | 0x27 |
| TAB | 0xf |
| CAPSLOCK | 0x3a |
| ENTER | 0x1c |
| LEFT_CONTROL | 0x1d |
| LEFT_ALT | 0x38 |
| LEFT_SHIFT | 0x2a |
| RIGHT_SHIFT | 0x36 |
| TILDE | 0x29 |
| PRINTSCREEN | 0x37 |
| NUM_LOCK | 0x45 |
| SPACE | 0x39 |
| DELETE | 0x53 |
| COMMA | 0x33 |
| PERIOD | 0x34 |
| BACKSLASH | 0x35 |
| FORWARDSLASH | 0x2b |
| LEFT_BRACKET | 0x1a |
| RIGHT_BRACKET | 0x1b |
| F1 | 0x3b |
| F2 | 0x3c |
| F3 | 0x3d |
| F4 | 0x3e |
| F5 | 0x3f |
| F6 | 0x40 |
| F7 | 0x41 |
| F8 | 0x42 |
| F9 | 0x43 |
| F10 | 0x44 |
| F11 | 0x57 |
| F12 | 0x58 |
| NUMPAD_0 | 0x52 |
| NUMPAD_1 | 0x4f |
| NUMPAD_2 | 0x50 |
| NUMPAD_3 | 0x51 |
| NUMPAD_4 | 0x4b |
| NUMPAD_5 | 0x4c |
| NUMPAD_6 | 0x4d |
| NUMPAD_7 | 0x47 |
| NUMPAD_8 | 0x48 |
| NUMPAD_9 | 0x49 |
| NUMPAD_PLUS | 0x4e |
| NUMPAD_MINUS | 0x4a |
| NUMPAD_PERIOD | 0x53 |
| NUMPAD_ENTER | 0x9c |
| NUMPAD_BACKSLASH | 0xb5 |
| LEFT_MOUSE | 0x100 |
| RIGHT_MOUSE | 0x101 |
| MIDDLE_MOUSE | 0x102 |
| MOUSE3 | 0x103 |
| MOUSE4 | 0x104 |
| MOUSE5 | 0x105 |
| MOUSE6 | 0x106 |
| MOUSE7 | 0x107 |
| MOUSE_WHEEL_UP | 0x108 |
| MOUSE_WHEEL_DOWN | 0x109 |