# This program send automatic emails
# T
# T

import smtplib
import pandas as pd
import os
import re
import  stdiomask as sm

# setting up a local smtplib server for testing
# in  order to run properly, it must be start as administrator previlages


def headerformatter(sent, wd=60, nl=[0,0]):
    if nl[0]==True and nl[1]==False:
        print("{:_^{:d}}".format("",wd))
        
    elif nl[0]==False and nl[1]==True:
        print("{:^}".format(""))
    
    elif nl[1]==True and nl[0]==True:
        print("\n{:_^{:d}}".format("",wd))
        print("{:}".format(""))
        
    elif nl[1]==False and nl[0]==False:
        print("\n{:_^{:d}}".format("",wd))
        print("\n{:_^{:d}}".format(sent,wd))
        print("{:_^{:d}}".format("",wd))
    
    else:
        print("\n{:_^{:d}}".format("",wd))
        print("\n{:_^{:d}}".format(sent,wd))
        print("{:_^{:d}}\n".format("",wd))
    
def readCsvTxtData(my_file="."):
    return pd.read_csv(my_file, header=0, encoding = 'unicode_escape',
                       sep='[\t|,]', engine='python')

def directoryFiles(fp="", ext="txt"):
    #list lis of files with the specified extension eg txt, jpg

    ext = tuple([ss.strip() for ss in (ext.upper().split(",") + ext.lower().split(","))])

    files = [] #holds file name
    file_path = [] #holds file path
    for file in os.listdir("./"+fp):
        if file.endswith(ext):
            files.append(file)
            file_path.append(os.path.join(fp, file))
            #print(os.path.join(file_path, file))

    for idx, val in enumerate(files,1):
        print("{:>6}\t\t{}".format(idx,val))
        
def listOptions(lst):
    '''
    the function takes a list and convert it into many choice option
    '''
    for k, e in enumerate(lst,1):
      print("{:^15}{:<10}".format(k,e))

def requestInput(st):
    '''
    for handling all inputs
    '''
    return input(st+": ")

def readEmailContent(file):
    """
    This reads email content as a single text

    Arguments:
        file {[txt file]} -- [contains the email content]

    Returns:
        [str] -- [returns the processed text file]
    """
    with open(file,mode="r",encoding="utf-8") as content:
        return content.read()

def validateEmail(email):
    
    """
    the function validate the email address supplied by the user

    Returns:
        [str] -- [returns true if email is valid and false if not not valid]
    """
  
    # Make a regular expression 
    regex = '^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'
    
    if(re.search(regex,email)):  
        return True 
    else:  
        print("\n**The specified email \""+str(email)+"\" is invalid. Please try again by re-entering the email ") 
        return False
    



headerformatter("  Program Usage Instructions  ", 86)
print('''
    ###############################################################################
    #                                                                             #
    #   The program automatically sends multiple independent emails and attaches   #
    #   certificate image file for the recipient.                                 #
    #                                                                             #
    #                                                                             #
    #   The program consists of two main directory                                #
    #       1. certs: this directory contains all the generate LT Certificate      #
    #       2. attendees: this directory contains csv or text file containing      #
    #          the 3 columns pointing to "firstname","last name or surname"       #
    #          and "email"                                                        #
    #                                                                             #
    ###############################################################################
      ''')
        
# servers
address =["smtp-mail.outlook.com", "smtp.mail.yahoo.com","smtp.yandex.com"]

server_name = ["outlook", "yahoo", "yandex"]

ports = [587,587, 587]

servers = {s+1 : [address[s].strip(), ports[s]] for s in range(len(server_name))}

# servers
headerformatter(" Select the email provider ", 86, [2,2])

#option selection
listOptions(server_name)
headerformatter(" w", 86, [1,0])
server_choice =  2#int(requestInput("Please select email provider number"))
#print(servers[server_choice][0],servers[server_choice][1])

#Enter your email
headerformatter("", 86, [1,0])

# while True:
    
#     sender_email = requestInput("Please enter your " + str(server_name[server_choice]) + " email address")
    
#     if validateEmail(sender_email) == True:
#         headerformatter("", 86, [1,0])
#         break
#     #print(sender_email)
#     headerformatter("", 86, [1,0])
    

sender_email =  ''
recipient_email = ''

# getting sender's password using getpass library
sender_password = ""#sm.getpass("Please enter your email password :", "*") #yahoo
# sender_password = "" #yandex



subject = "Leadership Training Certificate"
body = """This is a test e-mail message."""
msg = f'subject: {subject} \n\n {body}'

# try:
server_choice = smtplib.SMTP(servers[server_choice][0],servers[server_choice][1])
# server_choice = smtplib.SMTP("localhost",1025)
server_choice.set_debuglevel(1)
server_choice.ehlo()
server_choice.starttls()
server_choice.ehlo()    
server_choice.login(sender_email,sender_password)
server_choice.sendmail(sender_email, recipient_email,msg)
server_choice.quit()
print ("Successfully sent email")
# except Exception as e:
    # print (e)