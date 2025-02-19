import comtypes.client
import os
import logging
import re
import base64
from email.message import EmailMessage
import mimetypes

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

PP_FOLDER_PATH = os.path.abspath('PowerPoints')
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://mail.google.com/"]

def convert(inputFileName, outputFileName, formatType = 32):

    inputFileName = os.path.abspath(inputFileName)
    # checks if the inputFileName (pptx file) exists
    if not os.path.exists(inputFileName):
        logger.error("Input file does not exists.")  
        return

    # create an empty pdf file if the outputFileName does not exists
    if not os.path.exists(outputFileName):
        with open(outputFileName + '.pdf', 'w') as f: 
            f.write("")
        logger.info("Output file was created.")

    inputFileName = os.path.abspath(inputFileName)
    outputFileName = os.path.abspath(outputFileName)

    # create instance of powerpoint and open window
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    
    try:
        # open presentation and save it as pdf
        logger.info("Converting file.")
        deck = powerpoint.Presentations.Open(inputFileName)
        deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    except Exception as e:
        logger.error(f'{e}')
    finally:
        powerpoint.Quit()

    logger.info(f'Done.')

    return os.path.basename(outputFileName) # return the file name.

def get_files(path):
    
    files = [] # list to store all the files

    obj = os.scandir(path) # get all the files and dir of the folder

    for entry in obj:
        if entry.is_file(): # if entry is a file append to list of file
            files.append(entry.name)

    return files

def gmail_send_message(files, email):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("gmail", "v1", credentials=creds)
        message = EmailMessage()

        message.set_content("This is automated draft mail")

        message["To"] = email
        message["From"] = "joseangel130414@gmail.com"
        message["Subject"] = "IGNORE: PPTX to PYTHON"

        # loop through the list of pdf files
        for file in files:
            # get main and sub type = image/jpg (main/sub)
            type_subtype, _ = mimetypes.guess_type(file)
            maintype, subtype = type_subtype.split("/")

            # read file as binary and attach it to the message.
            with open(file, "rb") as fp:
                attachment_data = fp.read()
                message.add_attachment(attachment_data, maintype, subtype, filename=file)

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {"raw": encoded_message}
        # pylint: disable=E1101
        send_message = (
            service.users()
            .messages()
            .send(userId="me", body=create_message)
            .execute()
        )
        logger.info('Email has been sent.')
    except HttpError as error:
        print(f"An error occurred: {error}")
        logger.error(f'Error: {error}')
        send_message = None
    return send_message


files = get_files(PP_FOLDER_PATH) # path to where the powerpoints are saved

PDF_files = []

for file in files:
    file = os.path.join(PP_FOLDER_PATH, file)
    basename = os.path.basename(file) # get base name of the file
    file_name = re.match('^(.*)\.pptx*$', basename)
    file_output_name = file_name.group(1) # get the name of file without extension
    PDF_files.append(convert(file, file_output_name) + '.pdf') # convert file  

# send message
gmail_send_message(PDF_files, 'joseangel130414@gmail.com')
