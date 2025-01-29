import os
import subprocess
from openpyxl import load_workbook
from google.cloud import texttospeech

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service_account.json"
client = texttospeech.TextToSpeechClient()

workbook = load_workbook("vocab.xlsx")
sheet = workbook.active

os.makedirs("components", exist_ok=True)
os.makedirs("output", exist_ok=True)

max_rows = 3

columns_order = ["de", "ru", "b1_de", "b1_ru", "b2_de", "b2_ru"]

voice_map = {
    "de": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-B"),
    "ru": texttospeech.VoiceSelectionParams(language_code="ru-RU", name="ru-RU-Wavenet-D"),
    "b1_de": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-B"),
    "b1_ru": texttospeech.VoiceSelectionParams(language_code="ru-RU", name="ru-RU-Wavenet-D"),
    "b1_de_repeat": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-B"),
    "b2_de": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-C"),
    "b2_ru": texttospeech.VoiceSelectionParams(language_code="ru-RU", name="ru-RU-Wavenet-C"),
    "b2_de_repeat": texttospeech.VoiceSelectionParams(language_code="de-DE", name="de-DE-Studio-C")
}

audio_config = texttospeech.AudioConfig(
    audio_encoding=texttospeech.AudioEncoding.MP3,
    speaking_rate=1.0,
    pitch=1
)


def generate_speech(text, voice, filename):
    if not text or not text.strip():
        print(f"Skipping empty text for {filename}")
        return None

    path = f"components/{filename}"
    if os.path.exists(path):
        print(f"File already exists: {path}")
        return path

    print(f"Generating speech: {path}")
    synthesis_input = texttospeech.SynthesisInput(text=text)
    response = client.synthesize_speech(
        input=synthesis_input, voice=voice, audio_config=audio_config
    )

    with open(path, "wb") as f:
        f.write(response.audio_content)

    if not os.path.exists(path) or os.path.getsize(path) == 0:
        print(f"File was not created: {path}")
        return None

    return path


def generate_silence(duration_ms, filename):
    path = f"components/{filename}"
    if os.path.exists(path):
        return path
    subprocess.run(
        ["ffmpeg", "-f", "lavfi", "-i", "anullsrc=r=44100:cl=mono", "-t",
            str(duration_ms / 1000), "-q:a", "9", "-acodec", "mp3", path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    return path if os.path.exists(path) else None


one_sec_silence = generate_silence(1000, "silence_1s.mp3")
two_sec_silence = generate_silence(2000, "silence_2s.mp3")

final_audio_files = []

for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
    if row_idx > max_rows:
        break

    row_audio_files = []
    audio_paths = {}

    for col_idx, text in enumerate(row):
        if col_idx >= len(columns_order):
            break

        col_name = columns_order[col_idx]
        voice = voice_map.get(col_name)
        filename = f"{col_name}_{row_idx}.mp3"
        audio_path = generate_speech(text, voice, filename)

        if audio_path:
            audio_paths[col_name] = audio_path

    for col_name in columns_order:
        if col_name in audio_paths:
            row_audio_files.append(audio_paths[col_name])
            row_audio_files.append(one_sec_silence)

        if col_name == "b1_ru" and "b1_de" in audio_paths:
            row_audio_files.append(audio_paths["b1_de"])
            row_audio_files.append(one_sec_silence)

        if col_name == "b2_ru" and "b2_de" in audio_paths:
            row_audio_files.append(audio_paths["b2_de"])
            row_audio_files.append(one_sec_silence)

    if row_audio_files:
        row_audio_files.pop()
        row_output = f"output/row_{row_idx}.mp3"

        with open(f"components/row_{row_idx}.txt", "w") as f:
            for audio_file in row_audio_files:
                if not os.path.exists(audio_file):
                    print(f"Missing file {audio_file}")
                f.write(f"file '{os.path.abspath(audio_file)}'\n")

        subprocess.run(
            ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i",
                f"components/row_{row_idx}.txt", "-acodec", "mp3", row_output],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if not os.path.exists(row_output):
            print(f"Failed to create {row_output}")
        else:
            final_audio_files.append(row_output)
            final_audio_files.append(two_sec_silence)

if final_audio_files:
    final_audio_files.pop()
    final_audio_output = "output/final_audio.mp3"

    with open("components/final_list.txt", "w") as f:
        for audio_file in final_audio_files:
            f.write(f"file '{os.path.abspath(audio_file)}'\n")

    subprocess.run(
        ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i",
            "components/final_list.txt", "-acodec", "mp3", final_audio_output],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )

    if os.path.exists(final_audio_output):
        print(f"Final audiofile created: {final_audio_output}")
    else:
        print(f"Final audio file was not created")
