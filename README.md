### Setup .env

It should include the following:
```bash
CRED_NAME=client_secret_...
FOLDER_ID=1A6t....
FILE_ID=''  # Not necessary
CP_EXCEL_TO='' # Not necessary
```

### generate .exe file for Windows

The exe file should be compiled in a Windows environment.
To generate an Excel file, execute the .exe file in the same folder, specifying the cacert.pem file. 

$ pyinstaller -F parsing_and_upload_to_drive.py


### generate macOS executable file

$ pyinstaller --onefile --windowed --icon=./bagofmoney_5108.ico parsing_and_upload_to_drive.py
