
import pandas as pd
import pywhatkit
import time
from tkinter import Tk, filedialog
import tkinter.simpledialog as sd
import pyttsx3
import os

# --- Voice setup ---
engine = pyttsx3.init('sapi5')
engine.setProperty('rate', 160)
engine.setProperty('volume', 1)

def speak(text):
    engine.say(text)
    engine.runAndWait()

# --- Fancy Welcome Banner ---
def welcome_banner():
    os.system("cls" if os.name == "nt" else "clear")  # Clear console (Windows/Linux)
    print("="*60)
    print("🚀  Welcome to WhatsApp Sender".center(60))
    print("👨‍💻  Created by: Awais Ali (github.com/awaisali532)".center(60))
    print("="*60)
    print("\nThis program will help you send bulk WhatsApp messages 📩")
    print("Make sure your numbers are in an Excel file with column: Phone\n")
    input("👉 Press ENTER to select your Excel file... ")

welcome_banner()

# --- File selection dialog ---
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select Excel File with Phone Numbers",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not file_path:
    print("❌ No file selected, exiting...")
    exit()

# --- Read file ---
df = pd.read_excel(file_path, engine="openpyxl")

if 'Phone' not in df.columns:
    print("❌ Excel file must have a column named 'Phone'")
    exit()

if 'Status' not in df.columns:
    df['Status'] = ""

numbers = df['Phone'].tolist()

# --- Ask user for custom message ---
message = sd.askstring("Message Input", "Enter the WhatsApp message to send:")
if not message:
    print("❌ No message entered, exiting...")
    exit()

# Parameters
WAIT_TIME = 8
MAX_RETRIES = 2

# --- Announce start ---
print("\n▶️ Program starting to send messages...")
speak("Program starting to send messages")

try:
    for idx, number in enumerate(numbers):
        number_str = str(number).strip()

        # Skip already done numbers
        if str(df.at[idx, 'Status']).lower() == "done":
            print(f"⏭️ Skipping {number_str} (already done)")
            continue

        # --- Validation ---
        if not number_str or number_str.lower() == "nan":
            print(f"⚠️ Row {idx+1} skipped (empty number)")
            df.at[idx, 'Status'] = "Invalid/Empty"
            df.to_excel(file_path, index=False, engine="openpyxl")
            continue

        digits_only = ''.join(filter(str.isdigit, number_str))

        # Case 1: Already in +92 format
        if number_str.startswith("+92"):
            if len(digits_only) != 12:  # +92 + 10 digits
                print(f"⚠️ Skipping {number_str} (invalid +92 format)")
                df.at[idx, 'Status'] = "Invalid Number"
                df.to_excel(file_path, index=False, engine="openpyxl")
                continue
            final_number = number_str

        # Case 2: Local format starting with 03...
        elif digits_only.startswith("03") and len(digits_only) == 11:
            final_number = "+92" + digits_only[1:]  # remove '0' and add +92

        else:
            print(f"⚠️ Skipping {number_str} (not a valid PK number)")
            df.at[idx, 'Status'] = "Invalid Number"
            df.to_excel(file_path, index=False, engine="openpyxl")
            continue

        print(f"\n🟡 Sending to {final_number}...")
        retry_count = 0
        success = False

        while retry_count < MAX_RETRIES:
            try:
                pywhatkit.sendwhatmsg_instantly(
                    final_number,
                    message,
                    wait_time=WAIT_TIME,
                    tab_close=True
                )

                print(f"⌛ Waiting {WAIT_TIME} seconds...")
                time.sleep(WAIT_TIME + 2)

                print(f"✅ Sent to {final_number}")
                df.at[idx, 'Status'] = "Done"
                df.to_excel(file_path, index=False, engine="openpyxl")
                success = True
                break

            except Exception as e:
                retry_count += 1
                print(f"❌ Attempt {retry_count} failed for {final_number}: {e}")
                if retry_count >= MAX_RETRIES:
                    df.at[idx, 'Status'] = "Failed/Invalid"
                    df.to_excel(file_path, index=False, engine="openpyxl")
                time.sleep(5)

        if not success:
            print(f"❌ Message NOT sent to {final_number}")

        time.sleep(5)

finally:
    print("\n🎉 All tasks completed.")
    engine.say("All messages have been processed. Task completed")
    engine.runAndWait()
    engine.stop()
    time.sleep(1)

