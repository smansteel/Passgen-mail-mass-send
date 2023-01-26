import smtplib, ssl
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import os
import pandas as pd
sender_email = ''
password = ''
totalmail = []

file_name =  """FILEPATH with / instead of \ """# path to file + file name
sheet =  "Sheet1"# sheet name or sheet number or list of sheet numbers and names

df = pd.ExcelFile(file_name).parse(sheet) # Read from excel file


if input("Do you want to start the mass send from this excel file? y/[n]").lower() =="y":
    for i in range(0,len(df)):


        name_content = df.iat[i,1][2:]
        mdp_content = df.iat[i, 2]

        html = f"""\
        <html>
          <body class=MsoNormal>
          
          Some HTML formatted text
          
          </body>
        </html>
        """

        receiver_email = name_content 

        # Create MIMEMultipart object
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "" # Object
        msg["From"] = sender_email  #You can set something different if you want
        msg["To"] = receiver_email

        part = MIMEText(html, "html")
        msg.attach(part)
        encoders.encode_base64(part)
        msg.attach(part)
        # Create secure SMTP connection and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("Webmail server adres", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(
                sender_email, receiver_email, msg.as_string()
            )
        totalmail.append(receiver_email)
        print("sent successfully", name_content)

    html2 = f"""\
    <html>
      <body class=MsoNormal>
      
      Confirmation e-mail, set your text
     <hr>

      </body>
    </html>
    """


    html3 = f"""\
    <html>
      <body class=MsoNormal><span style='font-size:10.0pt;font-family:"Century Gothic",sans-serif;
    color:black;mso-fareast-language:FR'>
    
     <hr>

    <p >list of user that received an e-mail :</p>

    <p>{totalmail}</p>

      </body>
    </html>
    """
    receiver_email = 'SET YOUR CONFIRMATION E-MAIL RECEPIENT'

    # Create MIMEMultipart object
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Mass send"
    msg["From"] = sender_email
    msg["To"] = receiver_email

    part = MIMEText(html2+html+html3, "html")
    msg.attach(part)

    encoders.encode_base64(part)
    msg.attach(part)
    # Create secure SMTP connection and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("webmail.evok.ch", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, receiver_email, msg.as_string()
        )
    print("Mass send has finished")


else :
    print("Mass send canceled")
