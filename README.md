# email_data_table

I use easyimap to retrieve emails from the given email address and chosen parameters, then append each email's information regarding 'Sender', 'Title', 'Content', 'Date'.
to the eponymous lists.
Afterwards, I use pandas to convert these lists into DataFrame and also split the 'from_addr' attribute of easyimap into 'Name' and 'Address' respectively (the 'from_addr' is stored with the syntax 'Name <email>' by default). The DataFrame is then exported as an xlsx file. 
Everytime the code is run, an excel file containing the data of your email from newest to latest is generated. 
