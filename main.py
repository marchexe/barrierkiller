import os
from openpyxl import load_workbook
from google.cloud import texttospeech
from moviepy import TextClip, AudioFileClip, concatenate_videoclips
from moviepy.video.fx import CrossFadeIn, CrossFadeOut


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


def generate_video(text, duration):
    clip = TextClip(
        text=text,
        font="dejavu-sans-book.otf",
        font_size=60,
        color=(255, 255, 255),
        size=(1920, 1080),
        method="caption",
        text_align="center",
        bg_color=(30, 30, 30),
        duration=duration
    )

    return clip


video_clips = []

for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
    if row_idx > max_rows:
        break

    row_segments = []
    last_german_audio = None

    for col_idx, text in enumerate(row):
        if col_idx >= len(columns_order):
            break

        col_name = columns_order[col_idx]

        if "b2" in col_name:
            voice_key = "b2_de" if "de" in col_name else "b2_ru"
        else:
            voice_key = "de" if "de" in col_name else "ru"

        voice = voice_map[voice_key]
        filename = f"{col_name}_{row_idx}.mp3"
        audio = generate_speech(text, voice, filename)

        if audio:
            row_segments.append(audio)
            row_segments.append(AudioSegment.silent(duration=1000))

            if "b1_de" in col_name or "b2_de" in col_name:
                last_german_audio = audio

        if ("b1_ru" in col_name or "b2_ru" in col_name) and last_german_audio:
            row_segments.append(last_german_audio)
            row_segments.append(AudioSegment.silent(duration=1000))

    if row_segments:
        if len(row_segments) > 0:
            row_segments.pop()

        row_audio = sum(row_segments)

        audio_filename = f"output/audio_{row_idx}.mp3"
        row_audio.export(audio_filename, format="mp3")

        row_text = " | ".join([str(x) for x in row if x])

        duration = row_audio.duration_seconds

        video_clip = generate_video(row_text, duration)
        video_clip = video_clip.with_audio(AudioFileClip(audio_filename))

        video_clips.append(video_clip)

if video_clips:
    final_video = concatenate_videoclips(video_clips, method="chain")

    final_video.write_videofile(
        "output/final_video.mp4", fps=30, codec="libx264", audio_codec="aac")
    print("Final video file created: output/final_video.mp4")
else:
    print("No non-empty rows/cells were found for audio/video generation.")
