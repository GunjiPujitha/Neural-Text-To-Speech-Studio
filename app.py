import streamlit as st

from tts.engine import KOKORO_VOICE_CHOICES, VOICE_LABELS
from tts.kokoro_engine import generate_speech_with_kokoro
from tts.preprocess import preprocess_text


def main():
    st.set_page_config(page_title="Neural TTS Studio", page_icon="🎙️", layout="wide")

    st.markdown(
        """
        <style>
        .stApp {
            background: radial-gradient(circle at top left, #1c1f26 0%, #0e1117 45%, #07090d 100%);
            color: #e8ecf3;
        }
        .stTextArea textarea {
            background-color: #151922 !important;
            color: #eef2ff !important;
            border-radius: 12px !important;
            border: 1px solid #2c3345 !important;
        }
        .stButton button {
            background: linear-gradient(90deg, #0ea5e9, #2563eb);
            color: white;
            border-radius: 10px;
            height: 3em;
            font-weight: 700;
            border: none;
        }
        .stSelectbox div[data-baseweb="select"] > div,
        .stSlider div[data-baseweb="slider"] {
            background-color: #111622;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("🎙️ Neural Text-to-Speech Studio")
    st.caption("Fully offline, high-quality voice generation powered by Kokoro")

    text_input = st.text_area("Enter text", height=170, placeholder="Type something...")

    col1, col2 = st.columns(2)
    with col1:
        voice_key = st.selectbox(
            "Voice",
            KOKORO_VOICE_CHOICES,
            format_func=lambda v: VOICE_LABELS.get(v, v),
            index=0,
        )

    with col2:
        speed_value = st.slider("Speed", 0.5, 1.5, 1.0, step=0.05)

    if st.button("🎧 Generate Speech", use_container_width=True):
        if not text_input.strip():
            st.warning("Please enter text before generating speech.")
            return

        processed_text = preprocess_text(text_input)
        if not processed_text:
            st.warning("Processed text is empty. Please enter valid text.")
            return

        with st.spinner("Generating audio..."):
            try:
                audio_bytes = generate_speech_with_kokoro(
                    text=processed_text,
                    voice=voice_key,
                    speed=speed_value,
                )
            except Exception as error:
                st.error(f"Speech generation failed: {error}")
                return

        st.success("Speech generated successfully.")
        st.audio(audio_bytes, format="audio/wav")


if __name__ == "__main__":
    main()
