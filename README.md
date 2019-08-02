# Telegram printing bot
It's a simple Telegram bot. Long Polling method used for portability. You can send a document to it and if the document type is in accepted mime types bot can run a printing on a Windows machine default printer.
MSWord with command line arguments used for printing .docx .doc .txt formats.
And special silent utility PDFtoPrinter used for printing .pdf. Recommended format - PDF. You can specify your MSWord installation path with MICROSOFT_WORD_PATH constant inside bot.py
Tested with Windows 7 x86 + MSWord 2016.
This bot breathes a new life into old printers.
Also you need to specify your bot access token inside file named "credentials.json".
credentials.json format(JSON):
```
{
    "access_token": "string_from_bot_father"
}
```
How to run bot:
```
python3 bot.py
```
