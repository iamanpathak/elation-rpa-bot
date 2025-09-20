ElationBot 🚀

Smart RPA Automation for Elation EMR

ElationBot is a cross-platform automation tool designed to streamline repetitive workflows inside Elation EMR.
It intelligently combines web automation (Selenium), desktop automation (PyAutoGUI), and computer vision (OpenCV/Skimage + OCR) to handle tasks such as:

✨ Automated Login – Supports credentials + Google Authenticator (2FA).
✨ Patient Search – Retrieves patient data by name/DOB from Excel (OrderTemplate.xlsx).
✨ Document Uploads – Uploads signed order PDFs with intelligent drag-and-drop, OCR, and fallback detection.
✨ Smart Error Handling – Detects missing files, handles popups, and logs upload status to CSV.
✨ Batch Processing – Uploads files for multiple patients in one run.

⚙️ Key Features

🔑 Secure login with 2FA/Authenticator support

📂 Reads structured patient data from Excel + JSON configs

👓 Uses OCR + Template Matching to detect files and buttons

📊 Logs every upload in a CSV for audit tracking

🖥️ Works seamlessly across Windows, macOS, and Linux

🏁 Getting Started

Configure credentials & settings in config.json.

Place your OrderTemplate.xlsx and signed order PDFs in the correct folder.

Run the bot:

python main.py


Watch as it automates patient search & uploads 🔄
