
import imaplib
import excel
import providers
import email
import multiprocessing
import json
import concurrent.futures
from time import perf_counter, sleep
from collections import Counter

#global varibales
username = ""
apppassword = ""
gmail_host = 'imap.gmail.com'
mail = imaplib.IMAP4_SSL(gmail_host)
savedEmailCount = 0
newEmails = False

#reads data
def readData():
    global login_info, savedEmailCount,username
    with open('backupinformation.txt') as f:
        content = f.read()
        js = json.loads(content)
        login_info["apppassword"] = js["app_password"]
        savedEmailCount = js["saved_amount"]
        login_info["username"] = js["username"]


def sortEmail(chunk):
    global mail, login_info
    mail.login(login_info["username"],login_info["app_password"])
    mail.select("INBOX")
    partial_list = email_proccesor(chunk)
    mail.close()
    mail.logout()
    return partial_list


def email_proccesor(chunk):
    current_email = 0
    shared_list = []
    for num in chunk:

        _, data = mail.fetch(num, '(RFC822)')
        _, bytes_data = data[0]

        # convert the byte data to message
        email_message = email.message_from_bytes(bytes_data)
        if any(x["EMAIL"] == email_message["from"] for x in shared_list):
            for x in shared_list:
                if x["EMAIL"] == email_message["from"]:
                    current_email += 1
                    current_count = int(x["COUNT"])
                    new_count = 1 + current_count
                    x["COUNT"] = new_count
                    break
        else:
            current_email += 1
            provider = providers.Providers(email_message["from"], 1)
            shared_list.append(provider.getDict())
        print(
            f'Progress: {current_email} of {len(chunk)} Percent Complete: {round(current_email/len(chunk) * 100)}%', end='\r')
    return shared_list


def readEmail():
    global login_info, mail, shared_list, savedEmailCount, newEmails
    # select Inbox
    try:
        print(f'USERNAME: {login_info["username"]} App Password: {login_info["apppassword"]}')
        mail.login(login_info["username"], login_info["apppassword"])
        mail.select("INBOX")
        _, selected_mails = mail.search(None, 'ALL')
        print(
            f'\nTotal Messages in Inbox: {len(selected_mails[0].split())}\n\nREADING EMAILS NOW PLEASE WAIT!!!\n')
        print(f'SAVED EMAIL COUNT FROM LAST RUN: {savedEmailCount}')
        if(savedEmailCount - len(selected_mails[0].split()) > 25 or savedEmailCount == 0):
            # divide and conquer by using multiprocessing to speed up the search!!
            chunked_list = []
            savedEmailCount = {len(selected_mails[0].split())}
            newEmails = True
            # check if total amount of emails can be cleanly divided by 10 with no remainder.
            # if there is a remainder then round down aka "floor" the value and add 1
            chunk_size = len(selected_mails[0].split()) // 10 if len(
                selected_mails[0].split()) % 10 == 0 else len(selected_mails[0].split()) // 10 + 1

            # creating the seperate jobs by dividng it up into 10 seperate jobs
            mail.close()
            mail.logout()
            for i in range(0, len(selected_mails[0].split()), chunk_size):
                chunked_list.append(selected_mails[0].split()[i:i+chunk_size])
            start_time = perf_counter()
            chunked_shared_list = []
           
            with concurrent.futures.ProcessPoolExecutor() as executor:
                for result in executor.map(sortEmail, chunked_list):
                    print(f'length of the process dictionaries: {len(result)}')
                    for d in result:
                        chunked_shared_list.append(d)
            x = True
            i = 0
            print(f'LENGTH: {len(chunked_shared_list)} ')
            final_dict = Counter()
            for d in chunked_shared_list:
                final_dict[f'{d["EMAIL"]}'] += d["COUNT"]

            finish = perf_counter()
            shared_list = dict(final_dict)
            shared_list["app_password"] = apppassword
            shared_list["saved_amount"] = savedEmailCount
            with open('backupinformation.txt', 'w') as file:
                file.write(json.dumps(shared_list))
            print(
                f'Finished in all 10 proccesses in {round(finish - start_time, 2)} second(s)')
        else:
            print('hello')
   
    except imaplib.IMAP4.error as e:
        print(f'ERROR! CONNECTING WITH GMAIL: {e}')

    except FileNotFoundError as e:
        print(f'ERROR READING FILE NAME: {e}')
    except Exception as e:
        print(f'UNKNOWN ERROR OCCURED: {e}')



def main():
    global shared_list, excelProcess, newEmails
    # readpassword stored in apppassword.txt
    try:
        readData()
    except Exception as e:
        print(e)
    else:
    # read emails from gmail.com split the work between 10 proccessors
        readEmail()
    # write the final list to Excel senders.xlsx
        if newEmails == True:
            excelProcess.writeToExcel(shared_list)
    # open senders.xlsx
    # excel.Excel().openExcel()
    # # see if workbook is closed if so start reading flagged data otherwise wait
    # excel.Excel().readExcel()

    # excel.Excel().deleteEmails()

if __name__ == "__main__":
    manager = multiprocessing.Manager()
    login_info = manager.dict()
    shared_list = {}
    current_list = 0
    excelProcess = excel.Excel()
    main()
