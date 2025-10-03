import pandas as pd
import os
import time
from gtts import gTTS
import pygame

# ---------------- CONFIG ---------------- #
EXCEL_PATH = r"C:\Users\NXTWAVE\Downloads\Untitled spreadsheet.xlsx"
DEVICE_ID = "DURWQ45LGE9DQSW8"  # from adb devices

# ---------------- LOAD EXCEL ---------------- #
df = pd.read_excel(EXCEL_PATH)

required_cols = {"Name", "Presense", "Father"}
if not required_cols.issubset(df.columns):
    raise ValueError(f"Excel must contain columns: {required_cols}")

# ---------------- PROCESS STUDENTS ---------------- #
for _, row in df.iterrows():
    name = str(row["Name"]).strip()
    status = str(row["Presense"]).strip().lower()
    father_num = str(row["Father"]).strip()

    if status == "absent" and father_num:
        message = f"Your son {name} is not present in class today"
        print(f"[INFO] Calling {father_num} → {message}")

        # ⿡ Direct CALL
        os.system(f'adb -s {DEVICE_ID} shell am start -a android.intent.action.CALL -d tel:{father_num}')
        time.sleep(3)  # wait for call to connect

        # ⿢ Convert message to speech (TTS)
        filename = f"alert_{name}_{int(time.time())}.mp3"
        tts = gTTS(text=message, lang="en", slow=False)
        tts.save(filename)

        # ⿣ Play audio through laptop speakers (phone mic captures it)
        print("[INFO] Playing audio message...")
        pygame.mixer.init()
        pygame.mixer.music.load(filename)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            time.sleep(1)

        # ⿤ Open SMS app with same message prefilled
        print("[INFO] Sending SMS...")
        os.system(
            f'adb -s {DEVICE_ID} shell am start -a android.intent.action.SENDTO -d sms:{father_num} --es sms_body "{message}"'
        )

        print("[DONE] Alert sent to:", father_num)