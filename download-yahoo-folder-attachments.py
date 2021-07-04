import imaplib
import email
import os

email_address = "email@yahoo.co.in"
password = "password"
folder_to_fetch = "folder-name"
download_folder = "june-2021"
start_date = "1-Jun-2021"
end_date = "30-Jun-2021"
email_subject = "Daily Statement"

# Connect to an IMAP server
def connect(server, user, password):
    m = imaplib.IMAP4_SSL(server)
    m.login(user, password)
    return m


# Download all attachment files for a given email
def download_attachments_in_email(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1].decode('utf-8')
    mail = email.message_from_string(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))
            print("downloaded {}".format(part.get_filename()))


# Download all the attachment files for all emails in the inbox.
def download_yahoo_attachments(user, password, mailbox_name, outputdir, start_date, end_date, email_subject):
    os.mkdir(outputdir)
    c = connect("imap.mail.yahoo.com", user, password)
    try:
        _, _ = c.select('"{}"'.format(mailbox_name), readonly=True)
        resp, items = c.search(None, '(SUBJECT "{}" SINCE {} BEFORE {})'.format(email_subject, start_date, end_date))
        items = items[0].split()
        for email_id in items:
            download_attachments_in_email(c, email_id, outputdir)
    finally:
        c.logout()


download_yahoo_attachments(email_address, password, folder_to_fetch, download_folder, start_date, end_date, email_subject)
