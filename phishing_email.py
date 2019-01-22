import win32com.client, win32com, os, sys, time, requests, re, time
from urlscan import *

def email_check(reply, email):
    while True:
        print("\n\n" + reply.Body)
        try:
            user_in = int(input("1. Send\n2. Don't Send\n\n"))
            if user_in == 1:
                reply.Send()
                email.Categories = 'Green Category'
                break
            elif user_in == 2:
                email.UnRead = True
                break
            else:
                os.system("cls")
                print("Invalid Input")
        except Exception:
            os.system("cls")
            print("Invalid Input")

def url_scan(api_key, email_list):
    # for i in email_list:
        # params = {'apikey': api_key, 'url': email_list[2]}
        # response = requests.post('https://www.virustotal.com/vtapi/v2/url/scan', data=params)
        # json_response = response.json()
        # print(json_response)
        # print("\n")
        #
        # report_params = {'apikey': api_key, 'resource': email_list[2]}
        # report_response = requests.get('https://www.virustotal.com/vtapi/v2/url/report', params=report_params)
        # print(report_response.json())
    for i in email_list:
        phishing_email = UrlScan(api_key, i)
        phishing_email.submit()
        while True:
            try:
                phishing_email.checkStatus()
                scan_json = phishing_email.getJson()
                phishing_url = scan_json['page']['url']
                print(phishing_url)
                print('\n')
                break
            except Exception:
                print(Exception)
                time.sleep(10)

def main():
    api_key = ''            #Add your urlscan.io api key

    while True:
        os.system('cls')
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        phishing_inbox = outlook.Folders['phishing'].Folders['Inbox']
        messages = phishing_inbox.Items
        counter = 0
        max_mes = 0
        unread_list = []

        messages.Sort("[ReceivedTime]", True)

        for message in messages:
            if message.UnRead == True:
                unread_list.append(message)
                print("\t" + str(counter + 1) + ") " + str(unread_list[counter].Sender) + " - " + unread_list[counter].Subject + "\n")
                counter += 1
            elif max_mes == 50:
                break
            max_mes += 1
        if len(unread_list) >= 1:
            print('\nr - Refresh  q - Quit')
            selected_email = input('Select email by number: ')
            os.system('cls')
        else:
            print('No Unread Emails')
            selected_email = input('\nr - Refresh  q - Quit\n')
        try:
            if selected_email.lower() == 'r':
                continue
            elif selected_email.lower() == 'q':
                break
            elif int(selected_email) > 0 and int(selected_email) <= len(unread_list):
                unread_list[int(selected_email) - 1].UnRead = False
                print(unread_list[int(selected_email) - 1].Sender)
                print(unread_list[int(selected_email) - 1].Body)
                print("\n----------------------------------------------------------------\n")

                risky_url = re.findall('\<(.*?)\>',unread_list[int(selected_email) - 1].Body)                                   # Graps all the urls in an email
                risky_url = [i for i in risky_url if not re.match(r"[^@]+@[^@]+\.[^@]+", i)]

                url_scan(api_key, risky_url)

                user_in = input('1) Reply Spam\n2) Reply Duo\n3) Reply Postmaster\n4) Phishing Email\n5) Back Unread\n6) Back Read\n\nResponse: ')
            else:
                print('Invalid Input')
                time.sleep(1)
                continue
        except Exception:
            print('Invalid Input')
            time.sleep(1)
            continue

        if int(user_in) == 1:
            reply = unread_list[int(selected_email) - 1].Reply()
            reply.Body = "{},\n\nThank you for your Phishing e-mail submission.  After reviewing the e-mail, we have determined that this is not a Phishing attempt, but a form of SPAM mail.  To report SPAM mail, please use the option that is available on your Mimecast tab.   If this feature is not available in your Outlook, please submit a Remedy ticket to the Enterprise Architecture team so it can be added.\n\nThanks,\n\nMLH Information Security Team \n\n\n".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 2:
            reply = unread_list[int(selected_email) - 1].Reply()
            reply.Body = "{},\n\nThank you for being so vigilant with our Phishing awareness.  This e-mail is a legitimate e-mail from the IT Security team to enroll in our new two-factor authentication service, Duo, for remote users.  The link below explains Duo and the process. Again thank you for submitting your concerns.\n\nhttp://mlh.gomolli.org/about-us/non-clinical-departments/information-technology/DUO/index.dot\n\nMLH Information Security Team".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 3:
            reply = unread_list[int(selected_email) - 1].Reply()
            reply.Body = "{},\n\nThank you for being so vigilant with our Phishing awareness.  This e-mail is a legitimate e-mail from the IT department.  The “Message on hold” email is from Mimecast and it gives you a list of e-mail and gives you the option to “Release, Block, or Permit” the e-mail.  When reviewing theses message, please continue to be vigilant about which e-mail you release or permit to your inbox.  If you do not recognize the sender, feel free to block it and Mimecast will automatically block future messages from that sender.\n\nThanks,\n\nMLH Information Security Team".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 4:
            unread_list[int(selected_email) - 1].Categories = 'Yellow Category'

        elif int(user_in) == 5:
            unread_list[int(selected_email) - 1].UnRead = True

main()
