import numpy as np
import pandas as pd
import easyimap as ei
import datetime as dt

# Lists and parameters

sender = []
title = []
address = []
content = []
date = []

email_counts = 99 # the number of emails to retrieve to the data table. Enter n-1 for n emails

#Retrieving email contents and append to the lists

host = '' #for gmail use 'imap.gmail.com'
user = ''
password = '' # use 'App passwords' if use gmail

server = ei.connect(host,user,password)

for i in range(email_counts): # starting from i=0 is the newest one
    email = server.mail(server.listids(email_counts)[i]) # change 'listids' to 'unseen' if want to check only the uread emails
    sender.append(email.from_addr)
    title.append(email.title)
    content.append(email.body)
    date.append(email.date)

server.quit()

## Preparing DataFrame

# File name
now = dt.datetime.today().strftime('%Y-%m-%d %Hh%Mm%Ss') # stringformatting time to exclude microseconds and the ':'
filename = f'Inboxes of {user} as of {now}.xlsx'

# DF & add columns from the lists above
email_df = pd.DataFrame()

columns = [('Sender', sender), ('Address', address), ('Title', title), ('Content', content), ('Date',date)] 

for c in columns:
   email_df[c[0]] = pd.Series(c[1]) # iterate: email_df['Sender'] = pd.Series(sender)

# Splitting name and email
email_df.iloc[:,[0,1]] = email_df.iloc[:,0].str.split('<', expand=True)
email_df.iloc[:,1] = email_df.iloc[:,1].str.replace('>', '')

# Removing whitespace of the strings for filtering purpose
email_df = email_df.apply(lambda x: x.str.strip())
    # filtering by: email_df[email_df['Sender'] == 'sender_name']

#Optional: Calculating number of characters for each 'Content'
email_df['Content Characters'] = pd.Series([len(x) for x in email_df['Content']])

# Saving as excel files
email_df.to_excel(filename)


