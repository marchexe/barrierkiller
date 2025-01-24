import os

print("Generating audio...")
os.system("python ./py/audio.py")

print("Generating video..")
os.system("python ./py/video.py")

print("All done. Check path: 'output/final.mp4'")
