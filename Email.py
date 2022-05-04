from exchangelib import DELEGATE, Account, Credentials, Configuration, NTLM, Message, Mailbox, HTMLBody

def Email(to, subject, body):
    creds = Credentials(
        username=r'brt\notification',
        password='Start123!'
    )
    config = Configuration(server="owa.bitzer.cn",credentials=creds,auth_type=NTLM)
    account = Account(
        primary_smtp_address="notification@bitzer.cn",
        credentials=creds,
        config=config,
   #     autodiscover='True',
        access_type=DELEGATE
    )
    m = Message(
        account=account,
        subject=subject,
        body=HTMLBody(body),
        to_recipients=[Mailbox(email_address=to)]
    )

    m.send()
    return "success"
