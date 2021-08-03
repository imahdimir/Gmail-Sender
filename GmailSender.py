##
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from glob import glob
from pathlib import Path
import pandas as pd

if getattr(sys, 'frozen', False):
    cwd = Path(sys.executable).parent
elif __file__:
    cwd = Path(__file__).parent

def write_mail(sender, subj, index_mail, data_row):
    r = data_row

    msg = MIMEMultipart()

    msg['From'] = sender
    msg['To'] = index_mail
    msg['Subject'] = subj  # The subject line

    mail_cont = ''.join(
            str(col) + ':  ' + str(r[col]) + '\n' for col in r.index)

    msg.attach(MIMEText(mail_cont, 'plain'))
    return msg.as_string()

def main():
    pass
    ##
    with open(f'{cwd}/sender.txt', 'r') as f:
        gm_add_pass = f.read()

    sender_address = gm_add_pass.split('\n')[0]
    sender_pass = gm_add_pass.split('\n')[1]
    subject = gm_add_pass.split('\n')[2]

    xlpn = glob(f'{cwd}/*.xlsx')
    xlpn = [x for x in xlpn if x[:2] != '~$']
    assert len(xlpn) == 1, 'None or More than one Excel file'

    data = pd.read_excel(xlpn[0])
    data = data[data['Email'].notna()]
    data = data.set_index('Email')
    print(data)
    ##
    session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
    session.starttls()  # enable security
    session.login(sender_address, sender_pass)

    for ind, row in data.iterrows():
        text = write_mail(sender = sender_address,
                          subj = subject,
                          index_mail = ind,
                          data_row = row)
        session.sendmail(sender_address, ind, text)
        print(text)
        print(f'Mail Sent To {ind}')

    session.quit()

##
if __name__ == '__main__':
    main()
    print('Done!')

##
