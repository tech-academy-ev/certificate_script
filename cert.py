"""
Reads a csv File, performs a mail merge to create Certificates and then saves them as a PDF.
"""

from venv import create
from datetime import datetime
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
from docx2pdf import convert

def prepare_data(df):
    """
    Reads the data from a csv, then cleans and prepares it for the document creation process.
    :param df: The dataFrame that you want to prepare.
    :return: Returns a DataFrame
    """ 
    # Replace NaN Values with an empty string
    submissions_clean = df.fillna("")

    # Convert Kursnamen and Workshops from String to List
    submissions_clean['Kursnamen'] = submissions_clean['Kursnamen'].apply(lambda x: x[1:-1].split(',') if x != '' else x)
    submissions_clean['Workshops'] = submissions_clean['Workshops'].apply(lambda x: x[1:-1].split(',') if x != '' else x)

    submissions_pos = submissions_clean[submissions_clean['Pass / Failed'] == "Pass"].reset_index(drop=True)

    return submissions_pos


def create_certificate(submission):
    """
    Takes a single row from a DataFrame and creates a docx and pdf for it.
    :param submission: A single row from a DataFrame.
    :return: Returns True if everything worked fine.
    """ 
    # Import Template - For now only the Python Beginner Template
    template = DocxTemplate('Templates/'+ submission['Track'] + " " + submission['Level'] +'.docx')
    
    # Create Context for Word Doc
    context = { 
        'Name' : submission['Name'],
        'Vorname': submission['Vorname'],
        'Track': submission['Track'].replace('mit', 'with') if ('mit' in submission['Track']) else submission['Track'],
        'courses': submission['Kursnamen'],
        'workshops': submission['Workshops'],
        'Datum': datetime.today().strftime('%d.%m.%Y')
        }
        
    template.render(context)

    print("Saving Doc...")
    template.save("Certificates/"+ submission['Track'] +"/" + submission['Vorname'] + " " +submission['Nachname'] + " Certificate.docx")
    print("Doc saved.")
    # Zu lange (2-seitige) Zertifikate vielleicht in extra Ordner abspeichern für manuelle Korrektur?
    print("Converting docx to PDF...")
    convert("Certificates/" + submission['Track'] +"/" + submission['Vorname'] + " " +submission['Nachname'] + " Certificate.docx", "Certificates/"  + submission['Track'] + "/" + submission['Vorname'] + " " +submission['Nachname'] + " Certificate.pdf")
    print("PDF saved.")

    return True

#### Das Skript in ein Google Colab schieben. Und kommentieren, was angepasst werden muss. Zusätzlich read_excel zu read_csv einbauen.
def main():
    """
    Reads a csv File, performs a mail merge to create Certificates and then saves them as a PDF.
    """
    # Read CSV
    # submissions = pd.read_csv('Bewertungen.csv', delimiter=';') # Mit sep=None, engine='python' wird der erste Column Name um einen Char erweitert. Komischer Bug!
    submissions = pd.read_excel('Bewertungen.xlsx')

    submissions_prep = prepare_data(submissions)
    num_of_pos_submissions = str(submissions_prep.shape[0])

    print("Found " + num_of_pos_submissions + " positive submissions.")

    for index, submission in submissions_prep.iterrows():
        print("Creating Certificate " + str(index+1) + " out of " + num_of_pos_submissions + "...")
        
        # Create Certificates
        create_certificate(submission)
    
    print("+++ You did it! " + num_of_pos_submissions + " Certificates have been created. Time for some Chocolate! +++")
    print("Bye!")

if __name__ == '__main__':
    main()
