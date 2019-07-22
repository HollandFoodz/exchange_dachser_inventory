import smtplib, ssl

port = 587  # For starttls

def mail(server, sender, receiver, password, text):
    subject = 'Holland Foodz - import failed'
    message = 'Subject: {}\n\n{}'.format(subject, text)

    context = ssl.create_default_context()
    with smtplib.SMTP(server, port) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender, password)
        server.sendmail(sender, receiver, message)
