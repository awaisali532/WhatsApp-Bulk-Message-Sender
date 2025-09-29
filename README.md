**📱 WhatsApp Bulk Message Sender**

Automated Python tool to send bulk WhatsApp messages directly from an Excel sheet.
It leverages pandas, pywhatkit, tkinter, and pyttsx3 to make message sending simple and efficient.

**✨ Features **

📂 File Dialog → Select Excel file directly

📞 Phone Number Validation → Supports only +92 (PK format)

🗑️ Error Handling → Skips empty/invalid rows automatically

✅ Excel Updates → Marks Status (Done, Invalid Number, Failed)

🔁 Retries → Re-attempts sending on failure

🗣️ Voice Feedback → Announces start & completion via text-to-speech

🖥️ Console Output → Live progress tracking

**📋 Requirements**

Python 3.8+ installed

Install dependencies from requirements.txt:

pip install -r requirements.txt


**Dependencies used:**

pandas

pywhatkit

openpyxl

tkinter

pyttsx3

**🚀 How to Run**

**Clone this repository:**

git clone https://github.com/your-username/whatsapp-bulk-sender.git
cd whatsapp-bulk-sender


**Install dependencies:**

pip install -r requirements.txt


**Run the program:**

python whatsapp_sender.py

📂 Excel File Format

Your Excel file must have a column named Phone.
An additional column Status will be created automatically.

Phone	Status
+92 300 1234567	Done
+92 301 9876543	Failed
0300 7654321	Done
(empty)	Invalid/Empty
🛠️ Build as EXE (Optional)

**If you want to distribute as a standalone .exe:
**
pyinstaller --onefile --icon=logo.ico whatsapp_sender.py


--onefile → Generates a single exe

--icon=logo.ico → Add a custom logo (must be in .ico format)

🚧 Future Improvements

📊 GUI progress bar instead of console logs

🌍 Support for other country codes

🛡️ Better error handling for non-WhatsApp numbers

📜 License

This project is licensed under the MIT License.
