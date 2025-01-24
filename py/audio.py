import os
from pydub import AudioSegment
from openpyxl import load_workbook
from google.cloud import texttospeech

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'service_account.json'
client = texttospeech.TextToSpeechClient()

wb = load_workbook("vocab.xlsx")
ws = wb.active

os.makedirs('components', exist_ok=True)
os.makedirs('output', exist_ok=True)

max_rows = 2

columns_order = ["de", "ru", "b1_de", "b1_ru",
                 "b2_de", "b2_ru"]

voice_de = texttospeech.VoiceSelectionParams(
    language_code='de-DE', name='de-DE-Studio-B')
voice_ru = texttospeech.VoiceSelectionParams(
    language_code='ru-RU', name='ru-RU-Wavenet-D')

audio_config = texttospeech.AudioConfig(
    audio_encoding=texttospeech.AudioEncoding.MP3,
    speaking_rate=1.0,
    pitch=1
)


def generate_speech(text, voice, filename):
    if not text or not text.strip():
        return None
    
    path = f"components/{filename}"
    
    if os.path.exists(path):
        print(f"File already exists: {path}")
        return AudioSegment.from_file(path)
    
    print(f"Generating a new file: {path}")
    synthesis_input = texttospeech.SynthesisInput(text=text)
    response = client.synthesize_speech(
        input=synthesis_input, voice=voice, audio_config=audio_config)

    with open(path, "wb") as f:
        f.write(response.audio_content)
    return AudioSegment.from_file(path)


segments = []

for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
    if row_idx > max_rows:
        break

    row_segments = []

    for col_idx, text in enumerate(row):
        if col_idx >= len(columns_order):
            break

        col_name = columns_order[col_idx]
        voice = voice_de if "de" in col_name else voice_ru
        filename = f"{col_name}_{row_idx}.mp3"
        audio = generate_speech(text, voice, filename)

        if audio:
            row_segments.append(audio)
            row_segments.append(AudioSegment.silent(duration=1000))

    if row_segments:
        row_segments.pop()
        segments.extend(row_segments) 
        segments.append(AudioSegment.silent(duration=2000))


if segments:
    final_audio = sum(segments)
    final_audio.export("output/final.mp3", format="mp3")
    print("Audiofile created: output/final.mp3")
