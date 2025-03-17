# Voice Automation System

A Python-based voice control system that allows you to control your computer using voice commands.

## Features

- Voice recognition and text-to-speech feedback
- Open and close applications
- Web searching
- Wikipedia information lookup
- System information monitoring
- Volume control
- Screenshot capture
- Text typing
- And more!

## Installation

1. Make sure you have Python 3.7+ installed
2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the script:
```bash
python voice_assistant.py
```

2. Wait for the "Voice Assistant is ready" message

3. Speak your commands using the following format:
   - "open [application]" - Opens specified application (chrome, notepad, calculator)
   - "search [query]" - Searches Google for your query
   - "wikipedia [topic]" - Searches Wikipedia for information
   - "screenshot" - Takes a screenshot
   - "system" - Gets system information (CPU and memory usage)
   - "volume up/down/mute" - Controls system volume
   - "close [application]" - Closes specified application
   - "type [text]" - Types the specified text
   - "stop" or "exit" - Exits the program

## Examples

- "open chrome"
- "search latest news"
- "wikipedia artificial intelligence"
- "volume up"
- "screenshot"
- "system"
- "type hello world"

## Requirements

- Windows 10/11
- Python 3.7+
- Microphone
- Internet connection (for speech recognition and web features)

## Troubleshooting

1. If the voice recognition isn't working:
   - Check your microphone connection
   - Ensure you have a stable internet connection
   - Speak clearly and at a moderate pace

2. If applications won't open:
   - Verify the application is installed in the default location
   - Check if you have the necessary permissions

## Note

Some features may require administrative privileges to work properly. 