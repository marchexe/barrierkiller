import os
from openpyxl import load_workbook
from google.cloud import texttospeech



os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "/home/veritas/projects/private/barrierkiller/service-account.json"
client = texttospeech.TextToSpeechClient()

wb = load_workbook("vocab.xlsx")
ws = wb.active

os.makedirs("components", exist_ok=True)
os.makedirs("output", exist_ok=True)

max_rows = 2

columns_order = ["de", "ru", "b1_de", "b1_ru", "b2_de", "b2_ru"]

voice_map = {
    "de": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-B"),
    "ru": texttospeech.VoiceSelectionParams(language_code="ru-RU", name="ru-RU-Wavenet-D"),
    "b2_de": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-C"),
    "b2_ru": texttospeech.VoiceSelectionParams(language_code="ru-RU", name="ru-RU-Wavenet-C")
}

audio_config = texttospeech.AudioConfig(
audio_encoding=texttospeech.AudioEncoding.MP3,
    speaking_rate=1.0,
    pitch=1
)

input_text = texttospeech.SynthesisInput(text="Ich bin Barac O Bama ich liebe Democraty!")

response = client.synthesize_speech(input=input_text, voice=voice_map["de"],audio_config=audio_config)

with open("new.mp3", "wb") as out:
    out.write(response.audio_content)
    print("ready")
