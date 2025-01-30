import os
import subprocess
from openpyxl import load_workbook
from google.cloud import texttospeech
from PIL import Image, ImageDraw, ImageFont
from mutagen.mp3 import MP3
from moviepy import TextClip, AudioFileClip, concatenate_videoclips, vfx

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service_account.json"
client = texttospeech.TextToSpeechClient()

workbook = load_workbook("vocab.xlsx")
sheet = workbook.active

os.makedirs("components", exist_ok=True)
os.makedirs("output", exist_ok=True)
os.makedirs("slides", exist_ok=True)

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


def get_audio_duration(file_path):
    if not os.path.exists(file_path):
        return 1
    audio = MP3(file_path)
    return max(1, round(audio.info.length))


def generate_text_clip(text, duration, size=(1920, 1080)):
    text_clip = TextClip(
        font="dejavu-sans-book.otf",
        text=str(text),
        font_size=50,
        color=(255, 255, 255),
        bg_color=(30, 30, 30),
        text_align=center
    )

    text_clip = text_clip.with_duration(duration)
    text_clip = text_clip.with_position('center')

    text_clip = text_clip.with_effects(
        [vfx.CrossFadeIn(0.5), vfx.CrossFadeOut(0.5)])

    return text_clip


def get_clip_timing(audio_paths, silence_paths):
    timing = []
    current_time = 0

    for audio_path in audio_paths:
        if audio_path:
            duration = get_audio_duration(audio_path)
            timing.append((current_time, duration))
            current_time += duration

            silence_duration = get_audio_duration(silence_paths['one_sec'])
            current_time += silence_duration

        return timing


def generate_video(sheet, max_rows, columns_order, silence_paths):
    video_clips = []

    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if row_idx > max_rows:
            break

        row_audio_path = f"output/row_{row_idx}.mp3"
        if not os.path.exists(row_audio_path):
            continue

        audio_paths = []
        texts = []

        for col_idx, text in enumerate(row):
            if col_idx >= len(columns_order):
                break

            col_name = columns_order[col_idx]
            audio_path = f"components/{col_name}_{row_idx}.mp3"

            if os.path.exists(audio_path):
                audio_path.append(audio_path)
                texts.append(str(text))

                if col_name == "b1_ru":
                    repeat_path = f"components/b1_de_{row_idx}.mp3"
                    if os.path.exists(repeat_path):
                        audio_paths.append(repeat_path)
                        texts.append(str(row[columns_order.index("b1_de")]))
                    elif col_name == "b2_ru":
                        repeat_path = f"components/b2_de_{row_idx}.mp3"
                        if os.path.exists(repeat_path):
                            audio_paths.append(repeat_path)
                            texts.append(
                                str(row[columns_order.index("b2_de")]))

        clip_timing = get_clip_timing(audio_paths, silence_paths)

        for (text, (start_time, duration)) in zip(texts, clip_timing):
            if text and text.strip():
                clip = generate_text_clip(text, duration)
                video_clips.append(clip)

    if video_clips:
        try:
            final_video = concatenate_videoclips(video_clips, method=compose)

            final_video = AudioFileClip("output/final_audio.mp3")
            final_video = final_video.with_audio(final_audio)

            final_video.write_videofile(
                "output/final_video.mp4",
                fps=30,
                codec='libx264',
                audio_codec='aac'
            )

            final_video.close()
            final_audio.close()

            print("Video generation completed successfully")

        except Exception as e:
            print(f"Error during video generation: {str(e)}")

        finally:
            for clip in video_clips:
                clip.close()


if __name__ == "__main__":
    silence_paths = {
        'one_sec': generate_silence(1000, "silence_1s.mp3"),
        'two_sec': generate_silence(2000, "silence_2s.mp3")
    }

    if os.path.exists("output/final_audio.mp3"):
        generate_video(sheet, max_rows, columns_order, silence_paths)
