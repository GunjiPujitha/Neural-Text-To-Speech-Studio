# Text-to-Speech and Voice Generation using Machine Learning

A simple Streamlit application that converts text into speech using natural neural voices from Edge TTS with gTTS as a fallback option.

## Features

- Streamlit interface with text input and speech generation.
- Default high-quality neural voice support via Edge TTS.
- Voice selection for:
  - `en-IN-PrabhatNeural` (Indian Male)
  - `en-IN-NeerjaNeural` (Indian Female)
  - `en-US-JennyNeural` (US Female)
  - `en-GB-RyanNeural` (UK Male)
- Fallback online speech synthesis using gTTS.
- Audio playback directly in the browser.
- Improved text preprocessing for natural pauses and sentence flow.
- Edge speech controls: rate and pitch sliders.
- Enhanced Speech Mode and quick presets (Narration/Conversation/Announcement).
- Cached generation for repeated inputs.

## Installation

1. Clone or download the project.
2. Create a Python virtual environment.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run the application

```bash
streamlit run app.py
```

## Example usage

1. Open the Streamlit app in your browser.
2. Enter the text you want to convert to speech.
3. Choose `Neural (Edge TTS)` or `Standard (gTTS)`.
4. Select a voice if using Edge TTS.
5. Optionally tune `Speech Rate (%)` and `Pitch (Hz)` or choose a preset.
6. Click `Generate Speech`.
7. Listen to the audio directly in the browser.
