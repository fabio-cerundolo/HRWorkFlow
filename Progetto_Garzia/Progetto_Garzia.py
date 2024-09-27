import pandas as pd


#sopprime gli errori che riguardano le estensioni dei dati contenuti nei fogli excel
import warnings
warnings.simplefilter("ignore")

# input data del lavoro che vogliamo inviare
data_invio = input('Inserire la (dd/mm/yyyy) data dei lavori che si desiderano inviare: ')
from datetime import datetime
date = datetime.strptime(data_invio, "%d/%m/%Y").date()
if not date:
    print('inserire una data valida')
else:
    # apre foglio candidati
    file_path = "C:/Users/gamer/Documents/ACE10001I C-Lab HR/DB C-Lab (HRR).xlsx"
    foglio_candidati = pd.read_excel(file_path, sheet_name= 'Candidati')
    # localizziamo le righe che contengono la data
    posizione_data = foglio_candidati[foglio_candidati.eq(data_invio).any(axis=1)]
    # copiare le righe che ci interessano nel file da inviare
    righe_da_copiare = foglio_candidati.loc[posizione_data.index]
    new_file_path = "C:/Users/gamer/Documents/ACE10001I C-Lab HR/DB C-Lab_(Transfer).xlsx"
    righe_da_copiare.to_excel(new_file_path, index=False, sheet_name='Candidati')
    
    # cercare nel foglio anagskill i dati dei candidati selezionati
    foglio_anagskill = pd.read_excel(file_path, sheet_name= 'AnagSkill')
    anagskill_transfer = pd.DataFrame()
    nomi_curriculum = []
    for riga, row in righe_da_copiare.iterrows():
        identificativo_candidato = row['Id candidato']
        for y, q in foglio_anagskill.iterrows():
            valore = q['Progr Interno']
            if valore == identificativo_candidato:
                anagskill_transfer = pd.concat([anagskill_transfer, q], axis=1)
                cognome = q['Cognome']
                nome = q['Nome']
                cognome_nome = cognome + '_' + nome
                percorso_cv = 'C:/Users/gamer/Documents/ACE10001I C-Lab HR/CV al Cliente/cv_' + cognome_nome + '.pdf'
                print(percorso_cv)
                nomi_curriculum.append(percorso_cv)
    with pd.ExcelWriter(new_file_path, mode='a') as writer:
        anagskill_transfer.transpose().to_excel(writer, index=False, sheet_name= 'AnagSkill')


# automatizzazione mail
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

# Parametri di configurazione per l'invio della mail
smtp_server = 'smtp.office365.com'
smtp_port = 587
sender_email = 'indirizzo.HRR@outlook.it'
sender_password = 'indirizzoHRR'
receiver_email = 'centralesede29@gmail.com'
subject = 'Candidati'
body = """Alla cortese attenzione della Sede Centrale,
in allegato i file relativi ai candidati selezionati per le richieste ricevute.

Cordiali saluti"""

# Crea il messaggio
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = subject
message.attach(MIMEText(body, 'plain'))

# Funzione per allegare i file
def attach_file(filename):
    with open(filename, 'rb') as file:
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(file.read())
        attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(attachment)

# Allega i file
for nome in nomi_curriculum:       # personalizzare il nome del file cv in modo che alleghi il cv di ciascun candidato
    percorso_cv = 'C:/Users/gamer/Documents/ACE10001I C-Lab HR/CV al Cliente/cv_' + cognome_nome + '.pdf'
    attach_file(percorso_cv)


attach_file('C:/Users/gamer/Documents/ACE10001I C-Lab HR/DB C-Lab_(Transfer).xlsx')



try:
    # Connessione al server SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)

    # Invio della mail
    server.send_message(message)
    print("Mail inviata con successo!")

except Exception as e:
    print("Si è verificato un errore durante l'invio della mail:", str(e))

finally:
    # Chiusura della connessione
    server.quit()

