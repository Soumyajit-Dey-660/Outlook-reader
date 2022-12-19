from pathlib import Path  
import os
from win32com.client import Dispatch

# helper functions
def format_string_folder_name(subject_line):
    # characters_to_check = ['#', '%', '&', '{', '}', '<', '>', '*', '?','/','$','!',"'",':', '@','"','\\', '+', '`', '|', '=']
    characters_to_check = ['\\', '/', ':', '*', '?', '"', "<", ">", "|", " "]
    formatted_subject_line = ""

    for ch in subject_line:
        formatted_subject_line += ch if ch not in characters_to_check else "_"

    return formatted_subject_line

# Create output folder
output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Get all the subdirectories if Output folder already existsed
subfolders = [ f.path.split('\\')[-1] for f in os.scandir(Path.cwd() / "Output") if f.is_dir() ]

# Connect to outlook
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
#inbox = outlook.Folders("youremail@provider.com").Folders("Inbox")
inbox = outlook.GetDefaultFolder(6)
# https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# DeletedItems=3, Outbox=4, SentMail=5, Inbox=6, Drafts=16, FolderJunk=23

# Get messages
messages = inbox.Items

# message.subject
# message.senton               # return the date & time email sent
# message.senton.date()
# message.senton.time()
# message.sender
# message.SenderEmailAddress
# message.Attachments          # return all attachments in the email


for message in messages:
    try:
        subject = message.Subject
        body = message.body
        # attachments = message.Attachments

        # Limiting the number of character while craeting a filename to 50 characters
        # IMPORTANT: Tried with 100 characters but it gives too many IOError saying 
        # [Errno 2] No such file or directory:
        formatted_output_folder = format_string_folder_name(subject[:50])
        # If subfolder already exists, then create the folder by appending a the total number
        # of files that exist at the end. Otherwise create the folder as it is.
        existsing_same_folders_count = len([s for s in subfolders if formatted_output_folder in s])
        output_folder = formatted_output_folder \
            if existsing_same_folders_count == 0 \
            else f'{formatted_output_folder}_{existsing_same_folders_count}'
        

        # Create separate folder for each message
        target_folder = output_dir / str(output_folder)
        target_folder.mkdir(parents=True, exist_ok=True)

        # Add it to subfolders list
        subfolders.append(output_folder)

        # print(f'{target_folder}/EMAIL_BODY.txt')
        # break
        # Write body to text file
        with open(f'{target_folder}\\EMAIL_BODY.txt', "w", encoding="utf-8") as f:
            f.write(body)
        # Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

        # Save attachments
        # for attachment in attachments:
        #     attachment.SaveAsFile(target_folder / str(attachment))
    except IOError as e:
        print(f'IOError: {e}')
    except Exception as e:
        print(e)
        # continue
