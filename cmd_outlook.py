import win32com.client, win32com, os, sys, time
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

def main():
    while True:
        os.system('cls')
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        phishing_inbox = outlook.Folders['phishing'].Folders['Inbox']       #Change to whatever Inbox you want
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
            reply.Body = "{},\n\nSpam \n\n\n".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            # unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 2:
            reply = unread_list[int(selected_email) - 1].Reply()
            reply.Body = "{},\n\nDuo ".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            # unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 3:
            reply = unread_list[int(selected_email) - 1].Reply()
            reply.Body = "{},\n\nPostmaster ".format("Hey " + str(unread_list[int(selected_email) - 1].Sender)) #+ str(unread_list[int(selected_email) - 1].Body)
            # email_check(reply, unread_list[int(selected_email) - 1])
            reply.Send()
            # unread_list[int(selected_email) - 1].Categories = 'Green Category'

        elif int(user_in) == 4:
            # unread_list[int(selected_email) - 1].Categories = 'Yellow Category'

        elif int(user_in) == 5:
            unread_list[int(selected_email) - 1].UnRead = True

main()
