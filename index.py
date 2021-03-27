""" Modules nécessaires :
    - SpeechRecognition
    - Google Docs API
    - json
    - pyttsx3
    - os
    - time
    - pyfiglet
"""

#Import Google Docs API
from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

#Import SpeechRecognition Vocal
import speech_recognition as sr

#Voix robot
import pyttsx3

import json
import os
import pyfiglet
import time

class Program:

    # If modifying these scopes, delete the file token.json.
    SCOPES = None
    # The ID of a sample document.
    DOCUMENT_ID = None

    service = None
    document = None

    #Constructor
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/documents']
        self.DOCUMENT_ID = '1n_nUXgpOUnQwDvYa7QMK_bp-iWuclyKV6jE-rauhSSs'

    #Document configuration
    def configuration(self):
        creds = None

        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        self.service = build('docs', 'v1', credentials=creds)

        self.document = self.service.documents().get(documentId=self.DOCUMENT_ID).execute()
    
    #Main method
    def main(self):
        
        print(pyfiglet.figlet_format("DocsBot"))
        time.sleep(1)

        self.robot_say_something('Vous allez éditer le document "{}"'.format(self.document.get('title')))

        continue_program = True

        while continue_program:

            self.robot_say_something("Maintenant, que souhaitez vous faire sur votre document ? Vous avez le choix entre ajouter du texte, supprimer du texte ou ajouter une image")

            user_choice = input('ajouter du texte (1) / ajouter une image (2) / supprimer du texte (3) >> ')

            if user_choice == "1":
                self.insert_new_text()
            elif user_choice == "2":
                self.insert_new_image()
            elif user_choice == "3":
                self.delete_a_text()
            else:
                self.robot_say_something("Bon, vous ne souhaitez rien faire alors. Au revoir !")
                continue_program = False
        
    """ Methods for editing the document """
        
    #Method : the robot will say something
    def robot_say_something(self, speech):
        os.system('clear')
        engine = pyttsx3.init()
        print(speech)
        engine.say(speech)
        engine.runAndWait()

    #Method : dictate a text of the user
    def dictate_a_text(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            audio = r.listen(source)
        return r.recognize_google(audio, language="fr-FR")

    #Method : insert new text on the document   
    def insert_new_text(self):
        
        self.configuration()

        old_doc_index = self.document.get('body')["content"][-1]["endIndex"]

        self.robot_say_something("Dictez-moi votre texte.")

        if old_doc_index > 2:
            users_said = " " + self.dictate_a_text()
        else:
            users_said = self.dictate_a_text()

        requests = [
            {
                'insertText': {
                    'location': {
                        'index': old_doc_index - 1,
                    },
                    'text': users_said
                }
            }
        ]

        result = self.service.documents().batchUpdate(
            documentId=self.DOCUMENT_ID, body={'requests': requests}).execute()
        
        self.robot_say_something("Votre texte a bien été inséré dans le document Google Docs.")

    #Method : insert new image on the document   
    def insert_new_image(self):

        self.configuration()

        old_doc_index = self.document.get('body')["content"][-1]["endIndex"]

        time.sleep(10)

        self.robot_say_something("Je ne suis pas très intelligent donc veuillez m'indiquer l'adresse de votre image")

        image_adress = input(">> ")

        self.robot_say_something("Ok, je sais où elle se trouve ! Définissez la hauteur et la largeur de celle-ci")

        image_height = int(input('Hauteur >> '))
        image_width = int(input('Largeur >> '))

        requests = [
            {
                'insertInlineImage': {
                    'location': {
                        'index': old_doc_index - 1
                    },
                    'uri':
                        image_adress,
                    'objectSize': {
                        'height': {
                            'magnitude': image_height,
                            'unit': 'PT'
                        },
                        'width': {
                            'magnitude': image_width,
                            'unit': 'PT'
                        }
                    }
                }
            }
        ]

        # Execute the request.
        body = {'requests': requests}
        response = self.service.documents().batchUpdate(
            documentId=self.DOCUMENT_ID, body=body).execute()
        insert_inline_image_response = response.get('replies')[0].get(
            'insertInlineImage')
        
        self.robot_say_something("Votre image a bien été insérée dans le document Google Docs.")

    #Method : delete a text on the document
    def delete_a_text(self):

        self.configuration()

        full_dictionnary_docs_content = self.document.get('body')["content"]
        
        self.robot_say_something("Quel est le texte à supprimer ?")
        string_to_delete = self.dictate_a_text()

        is_the_text_deleted = False

        for content in full_dictionnary_docs_content[1:]: #[1:] tell the script to skip the first element in this list
            text_of_this_content = content["paragraph"]["elements"]#[0]["textRun"]["content"]
            for each_text in text_of_this_content:

                if "textRun" in each_text:
                    each_text_content = each_text["textRun"]["content"]

                    if string_to_delete in each_text_content:
                        new_string = each_text_content.split(string_to_delete)

                        length_of_the_string_to_delete = len(string_to_delete)
                        length_of_the_string_before_the_string_to_delete = len(new_string[0])

                        start_index = each_text["startIndex"] + length_of_the_string_before_the_string_to_delete
                        end_index = each_text["startIndex"] + (length_of_the_string_to_delete + length_of_the_string_before_the_string_to_delete)

                        requests = [
                            {
                                'deleteContentRange': {
                                    'range': {
                                        'startIndex': start_index,
                                        'endIndex': end_index,
                                    }
                                }
                            },
                        ]

                        result = self.service.documents().batchUpdate(
                            documentId=self.DOCUMENT_ID, body={'requests': requests}).execute()
                        
                        is_the_text_deleted = True
        
        if is_the_text_deleted:
            self.robot_say_something("Votre texte a bien été supprimé du document Google Docs.")
        else:
            self.robot_say_something("La chaîne à supprimer n'est pas dans le document.")
        
#Execution
if __name__ == '__main__':
    app = Program()
    app.configuration()
    app.main()
