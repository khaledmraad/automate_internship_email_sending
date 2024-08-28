import subprocess
import re
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def convert_latex_to_pdf(latex_file_path):
            try:
                subprocess.check_call(['latexmk', '-pdf', '-interaction=nonstopmode', latex_file_path])
                print("PDF generated successfully!")
            except subprocess.CalledProcessError as e:
                print("PDF generation failed. Error:", e)


def send_email(sender_name,sender_email, sender_password, receiver_email, subject, message, attachment_paths):
    # Create a multipart message object
    msg = MIMEMultipart()
    msg['From'] = f'{sender_name} <{sender_email}>' 
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Attach the message to the multipart object
    msg.attach(MIMEText(message, 'plain'))

    # Attach the PDF files
    for attachment_path in attachment_paths:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {attachment_path}')
            msg.attach(part)

    # Establish a secure connection with the SMTP server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        # Login to your Gmail account
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, receiver_email, msg.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print("Error sending email:", str(e))
    finally:
        # Close the connection
        server.quit()


# Specify the path to your CSV file
csv_file_path = "best401.csv"

# Open the CSV file
with open(csv_file_path, 'r', encoding='utf-8') as file:
    # Create a CSV reader object
    csv_reader = csv.reader(file)

    i=0
    # Iterate over each row in the CSV file
    for row in csv_reader:
        i+=1
        # Access individual columns using index
        column1 = row[0]
        column2 = row[1]
        column3 = row[2]

        print("-------------------------------------------------------")

        print(column1,column2,column3)
        print("-------------------------------------------------------")


        # Read the LaTeX file
        with open('template.tex', 'r') as file:
            latex_content = file.read()

        # Perform the word replacement using regular expressions
        old_word = r'charika'
        new_word = r'{}'.format(column1)
        modified_text = re.sub(r'\b' + re.escape(old_word) + r'\b', new_word, latex_content)

        old_word = r'blassa'
        new_word = r'{}'.format(column2)
        modified_text = re.sub(r'\b' + re.escape(old_word) + r'\b', new_word, modified_text)


        # Write the modified content back to the LaTeX file
        with open('Lettre de motivation.tex'.format(i), 'w') as file:
            file.write(modified_text)

        # Example usage
        latex_file_path = 'Lettre de motivation.tex'.format(i)
        convert_latex_to_pdf(latex_file_path)
        
        sender_email = "khaled.mrad@ensi-uma.tn"
        sender_password = "your gmail api key"
        receiver_email = column3

        
        subject = "Candidature pour un stage d'été chez {}".format(column1)

        message = """
Mes respectueuses salutations

J'espère que vous vous portez bien.

Je vous écris pour exprimer mon vif intérêt à obtenir un stage au sein de votre entreprise {}.

Je suis actuellement étudiant en deuxième année à l'École Nationale des Sciences de l'Informatique (ENSI). Je suis à la recherche d'un stage d'été d'une durée idéale de 6 à 8 semaines. 

Veuillez trouver ci-joint ma lettre de motivation et mon curriculum vitae (CV) pour votre considération.

Je vous remercie de prendre en considération ma candidature. J'attends avec impatience la possibilité de discuter de la manière dont mes compétences et mon enthousiasme peuvent contribuer au succès de {} . Si vous avez besoin de plus d'informations ou de documents supplémentaires, n'hésitez pas à me contacter. Je suis disponible à votre convenance pour un entretien.

Cordialement,
Mrad Khaled """.format(column1,column1)
        attachment_path = ["CVang.pdf","CVfr.pdf",'Lettre de motivation.pdf'.format(i)]

        # Send the email with the PDF attachment
        send_email("Khaled Mrad",sender_email, sender_password, receiver_email, subject, message, attachment_path)





                