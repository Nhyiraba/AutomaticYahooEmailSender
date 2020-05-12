# This program send automatic emails

import smtplib
import pandas as pd
import os
import re
# import  stdiomask as sm #used for masking input form user
from email.message import EmailMessage
import imghdr

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

def directoryFiles(filepath="", ext="txt"):
    #list lis of files with the specified extension eg txt, jpg

    ext = tuple([ss.strip() for ss in (ext.upper().split(",") + ext.lower().split(","))])

    files = [] #holds file name
    each_file_path = [] #holds file path
    for file in os.listdir(filepath):
        if file.endswith(ext):
            files.append(os.path.splitext(os.path.basename(file))[0])
            each_file_path.append(os.path.join(filepath, file))
            #print(os.path.join(fpath, file))

    for idx, val in enumerate(files,1):
        print("{:>6}\t\t{}".format(idx,val))

    return files, each_file_path

def columnBinder(dataSet, colIndex):
    
    if len(colIndex) == 0:
        return dataSet[dataSet.columns].apply(lambda x:
            ' '.join(x.dropna().astype(str)),axis=1 )
    else:
        col = [int(dd)-1 for dd in colIndex.split(',') ]
        return dataSet[dataSet.columns[col]].apply(lambda x:
            ' '.join(x.dropna().astype(str)),axis=1 )
      
def listOptions(lst):
    """the function takes a list and convert it into many choice option."""
    for k, e in enumerate(lst,1):
      print("{:^15}{:<10}".format(k,e))

def requestInput(st):
    """for handling all inputs."""
    return input(st+": ")

def readContent(file):
    """This reads email content as a single text.

    Arguments:
        file {[txt file]} -- [contains the email content]
    """
    
    with  open(file, "r", encoding = "utf-8") as f:
        return f.read()

def validateEmail(email):
    
    """the function validate the email address supplied by the user.

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
address =["smtp.mail.yahoo.com"]
server_name = ["yahoo"]
ports = [587]

servers = {s+1 : [address[s].strip(), ports[s]] for s in range(len(server_name))}

# servers
headerformatter(" Select the email provider ", 86, [2,2])

#option selection
listOptions(server_name)
headerformatter(" w", 86, [1,0])
server_choice =  1#int(requestInput("Please select email provider number"))


#Enter your email
headerformatter("", 86, [1,0])


#creating message object and contend
sender_email = ''
sender_password = "" #yahoo app generated password

# body = readContent(".\emailcontent.txt")
# body = body.encode('ascii', 'ignore').decode('ascii')
# subject = "Leadership Training Certificate"
# msg = f'Subject : {subject} \n\n {body}'

msg = EmailMessage()
body = readContent(".\emailcontent.txt")
body = body.encode('ascii', 'ignore').decode('ascii')
msg["subject"] = "Leadership Training Certificate"
msg["from"] = sender_email


# this section of the code will compare attendees nanme with certificate's file name
cert_files = directoryFiles("certs","jpg, jpeg, png") #this returns both directory and file name
att_data = readCsvTxtData("attendees/testDataSet.csv")
full_name = columnBinder(att_data,"2,1")
att_email = {full_name[i].strip() : att_data[att_data.columns[2]][i].strip() for i in range(len(full_name))}
att_cert = {cert_files[0][i].strip() : cert_files[1][i].strip() for i in range(len(cert_files[0]))}


for att in full_name:
    if att in att_cert:
        
        recipient_email = att_email[att]#"pastordan1991@outlook.com"
        msg["to"] = recipient_email
        msg.set_content(body)
        
        with open(att_cert[att], "rb") as cert_img:
            file_data = cert_img.read()
            file_type = imghdr.what(cert_img.name)
            file_name = cert_img.name
        
        msg.add_attachment(file_data, maintype="image", subtype=file_type, filename = file_name)
        

        try:
            server_choice = smtplib.SMTP(servers[server_choice][0],servers[server_choice][1])
            # server_choice = smtplib.SMTP("localhost",1025)
            server_choice.set_debuglevel(1)
            server_choice.ehlo()
            server_choice.starttls()
            server_choice.ehlo()    
            server_choice.login(sender_email,sender_password)
            #server_choice.sendmail(sender_email, recipient_email,msg)
            server_choice.send_message(msg)
            
            headerformatter("", 86, [1,0])
            print ("\nSuccessfully sent email\n")
            headerformatter("", 86, [1,0])
            
            server_choice.quit()
        except Exception as e:
            print (e)

        
            
    


#localhost tets run server
# python -m smtpd -c DebuggingServer -n local:1024
# this code can be be used to try localhost test and emailwill be
# displayed in the console

# try:
#     # server_choice = smtplib.SMTP("localhost",1025)
#     server_choice.sendmail(sender_email, recipient_email,msg)
#     server_choice.quit()
#     print ("Successfully sent email")
# except Exception as e:
#     print ("uanble to send email")
# ###########################

#sender email validation
# while True:
    
#     sender_email = requestInput("Please enter your " + str(server_name[server_choice]) + " email address")
    
#     if validateEmail(sender_email) == True:
#         headerformatter("", 86, [1,0])
#         break
#     #print(sender_email)
#     headerformatter("", 86, [1,0])
