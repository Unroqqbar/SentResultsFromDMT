import docx
import re
from redmail import gmail
from time import sleep
from tkinter import filedialog
import os
from dotenv import load_dotenv

load_dotenv(r"C:\Users\Steve\Dropbox\Python\Variables\.env.txt")
USER = os.getenv("GmailUsername")
PASSWORD = os.getenv("GmailPassword")


def extract_emails_from_paragraph(paragraph):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_pattern, paragraph.text)
    if emails:
        return emails
    else:
        return None


def create_email_dictionary(doc_path):
    doc = docx.Document(doc_path)
    emails_found_in_word = {}

    for paragraph in doc.paragraphs:
        # Extract the first two words in lower case
        words = paragraph.text.split()[:2]
        name_of_kid = ''.join(words).lower()
        name_of_kid = name_of_kid+".pdf"
        
        # Extract emails from the paragraph
        parent_email = extract_emails_from_paragraph(paragraph)

        # Add to the dictionary, if not already inside, and if email exists in a row from the Word document
        if name_of_kid not in emails_found_in_word and parent_email:
            emails_found_in_word[name_of_kid] = parent_email
        
    return emails_found_in_word


word_document_path = 'C:/Users/Steve/Dropbox/Python/SentPDFbyEmail/directory.docx'
email_dict = create_email_dictionary(word_document_path)

print('\nHere are all the names and corresponding E-mails from the directory.docx:')
for kid_name_from_word, email_from_word in email_dict.items():
    print(f'{kid_name_from_word}: {email_from_word}')
print('\n')

pdf_folder = filedialog.askdirectory(title="Select Folder with Pdf Files")
print(pdf_folder)

if pdf_folder != "":
    for kid_name, email in email_dict.items():
        gmail.username = USER
        gmail.password = PASSWORD

        try:
            gmail.send(
                subject="Deutscher Motorik Test",
                receivers=email,
                text="Chers parents,\n"
                     "en annexe, vous trouvez les résultats du Deutscher Motorik Test de votre enfant. \n"
                     "Si vous avez encore des questions, n'hésitez pas à nous contacter par fitkanner@mersch.lu\n\n"
                     "ATTENTION: Ne répondez pas à ce message!\n"
                     "Ce message a été créé automatiquement pour envoyer les résultats du Deutscher Motorik Test.\n"
                     "Toutes les communications continueront à passer par fitkanner@mersch.lu\n\n"
                     "cordialement",
                attachments={f"{pdf_folder}/{kid_name}"}
            )
        except ValueError:
            print(f"ERROR: The {kid_name} does not exist\n")
            print(f"{pdf_folder}/{kid_name}")
            sleep(1)
        else:
            print(f"The {kid_name} has been sent to {email}\n")
            sleep(3)
else:
    print("No Folder selected")
