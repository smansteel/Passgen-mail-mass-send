# Passgen-mail-mass-send

Here is :

- a password generator that takes for input an Excel file, add a column a create a password for each line

- a mass mailer, which takes for input an excel file too and sends an email to the email found in column 1 and can incorporate text from other column. The text of the e-mail can be written using html.


Required modules :  Pandas


Known bugs: 
- The first row of the excel file is considered a header so is not treated by neither the passgen or the massmailer
- If the emails that you send are to an external server, you might be blocked after a few emails.
- The password has to be in clear text in the code

If you have any question about this do not hesitate to contact me here or on discord : sman_steel#4595
