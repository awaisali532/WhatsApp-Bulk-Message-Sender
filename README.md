📱 WhatsApp Bulk Message Sender

Automated tool to send bulk WhatsApp messages directly from an Excel sheet.
It uses Python + pywhatkit + pandas + tkinter to simplify sending messages.

✨ Features

📂 Select Excel file via file dialog

📞 Validates phone numbers (+92 PK format only)

🗑️ Skips empty/invalid rows automatically

✅ Updates Excel with message Status (Done, Invalid Number, Failed)

🔁 Retries sending messages if any attempt fails

🗣️ Voice feedback using pyttsx3 (start & completion notifications)

🖥️ Console output for live tracking of progress

📋 Requirements

Make sure you have Python 3.8+ installed.
Install dependencies from requirements.txt:

pip install -r requirements.txt

Dependencies used:

pandas

pywhatkit

openpyxl

tkinter

pyttsx3

🚀 How to Run

Clone this repository:

git clone https://github.com/your-username/whatsapp-bulk-sender.git
cd whatsapp-bulk-sender

Install required dependencies:

pip install -r requirements.txt

Run the program:

python whatsapp_sender.py

📂 Excel File Format

Your Excel file must have a column named Phone.
An additional column Status will be created automatically.

Phone Status
+92 300 1234567 Done
+92 301 9876543 Failed
0300 7654321 Done
(empty) Invalid/Empty
🛠️ Build as EXE (Optional)

If you want to distribute as an .exe:

pyinstaller --onefile --icon=logo.ico whatsapp_sender.py

--onefile → Single exe file

--icon=logo.ico → Custom logo (must be .ico format)

🚧 Future Improvements

GUI progress bar instead of console logs

Support for other country codes

Better error handling for non-WhatsApp numbers

📜 License

This project is licensed under the MIT License.
