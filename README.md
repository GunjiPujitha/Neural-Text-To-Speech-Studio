The scope of Neural Text-to-Speech Studio is defined as follows. The system accepts plain
English text as input and produces high-quality WAV audio as output. It supports nine distinct voice
identities across American English (five voices) and British English (four voices), covering both
male and female speakers. The user can control the speech rate via a speed parameter ranging from
0.5× (half speed) to 1.5× (one-and-a-half times normal speed).
The system includes a text preprocessing module that handles common input quality issues:
numeric digit expansion, abbreviation substitution, punctuation normalisation, illegal character
removal, and sentence length management. The user interface is a single-page web application served
locally via the Streamlit framework and accessible from any web browser on the same machine.
The following capabilities are explicitly outside the scope of this version of the project:
multilingual synthesis (languages other than English), real-time streaming synthesis, voice cloning
from user-provided audio, batch file processing, persistent history or user accounts, and deployment
to a public cloud or multi-user server environment.
# Neural-Text-To-Speech-Studio
This repository is for practicing the GitHub Flow
