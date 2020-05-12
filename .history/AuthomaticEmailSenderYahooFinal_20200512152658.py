# This program send automatic emails

import smtplib
import pandas as pd
import os
import re
import imghdr

# import the corresponding modules
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# setting up a local smtplib server for testing
# in  order to run properly, it must be start as administrator previlages


def headerformatter(sent, wd=60, nl=[0,0]):
    
    """[summary]
    """    
    
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
    
    try:
        return pd.read_csv(my_file, header=0, encoding = 'unicode_escape',
                       sep='[\t|,]', engine='python')
    
    except Exception:
        return pd.read_excel(my_file, header=0, sheet_name=0)
    
    except Exception as e:
        print(e)


def removeSpace(dataSet):
    
    for sp in range(len(dataSet.columns)-1):
        
        
        dataSet[dataSet.columns[sp]] = dataSet[dataSet.columns[sp]].str.strip()
        dataSet[dataSet.columns[sp]] = dataSet[dataSet.columns[sp]].apply(lambda x: x.strip())
        dataSet[dataSet.columns[sp]] = dataSet[dataSet.columns[sp]].str.replace('\s{2,}', ' ')
        
    return dataSet

       

def directoryFiles(filepath="", ext="txt", f_d = True):
    #list lis of files with the specified extension eg txt, jpg
 

    ext = tuple([ss.strip() for ss in (ext.upper().split(",") + ext.lower().split(","))])

    files = [] #holds file name
    each_file_path = [] #holds file path
    for file in os.listdir(filepath):
        if file.endswith(ext):
            files.append(os.path.splitext(os.path.basename(file))[0])
            each_file_path.append(os.path.join(filepath, file))
            #print(os.path.join(fpath, file))
            
    if f_d == 1: 
        for idx, val in enumerate(files,1):
            print("{:>6}\t\t{}".format(idx,val))

    return files, each_file_path

def columnBinder(dataSet, colIndex):
    
    """this function conbines two or more columns together
    the user specifies the column numbers assigned to each column
    
    Dataset : pandas dataframe or table with columns
    colIndex : str pf int; eg:"1,2,3" 

    Returns:
        dataframe of combined columns
    """    
    
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
    #   The program automatically sends multiple independent emails and attaches  #
    #   certificate image file for the recipient.                                 #
    #                                                                             #
    #                                                                             #
    #   The program consists of two main directory                                #
    #       1. certs: this directory contains all the generate LT Certificate     #
    #       2. attendees: this directory contains csv or text file containing     #
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
smtp_server, port = servers[server_choice][0], servers[server_choice][1]

#Enter your email
headerformatter("", 86, [1,0])


#creating message object and contend
sender_email = ''
sender_password = "" #yahoo app generated password


msg = MIMEMultipart()
body = readContent(".\emailcontent.txt")
body = body.encode('ascii', 'ignore').decode('ascii')



# this section of the code will compare attendees nanme with certificate's file name
cert_files = directoryFiles("certs","jpg, jpeg, png") #this returns both directory and file name


data1 = readCsvTxtData("attendees/LTCertData1.lsx.xlsx")
data2 = readCsvTxtData("attendees/LTCertData2.lsx.xlsx")


att_data = removeSpace(pd.concat([data1, data2])).drop_duplicates()
att_data.to_excel(r'cleanedDataAtt2020Spring.xlsx', index = False, header=True)

att_new = readCsvTxtData('cleanedDataAtt2020Spring.xlsx')




##################################################

colNameIndex = [lst for lst in enumerate(att_new.columns,1)]


headerformatter("Enter the associated column number for sername and firstname by selecting the \nappropriate number next to the desired column by enter firstname followed by surname, \nthe order selected will appear oncertificates. eg (1, country), (2, sername), \n(3, firstname) enter --> 1,3 ",88)

print("{:30}".format("\nThe List columns in the loaded dataset\n"))
print("{:}".format(str(colNameIndex)))
headerformatter("",nl=[1,1],wd=85)

name_colIndex = input("Enter the respective columns number for name : " )
email_colIndex = input("Enter the respective columns for email : " )


full_name = [ names.strip() for names in columnBinder(att_new,name_colIndex)]
email = [em.strip() for em in columnBinder(att_new,email_colIndex)]

att_email = {full_name[i].strip() : email[i].strip() for i in range(len(full_name))}

cert_files = directoryFiles("certs","jpg, jpeg, png") #this returns both directory and file name
att_cert = {cert_files[0][i].strip() : cert_files[1][i].strip() for i in range(len(cert_files[0]))}



with smtplib.SMTP(smtp_server, port) as server:
    server.starttls()
    server.ehlo()
    server.login(sender_email, sender_password)
     
    #reader = pd.read_csv("attendees/testDataSet.csv")#csv.reader(file) reader[reader.columns[0]], reader[reader.columns[2]]
    for name in full_name:
        
        msg["subject"] = "Leadership Training Certificate"
        msg["from"] = sender_email  
        msg.attach(MIMEText(body, "plain"))
        
        
        with open(att_cert[name], "rb") as attachment:
            # The content type "application/octet-stream" means that a MIME attachment is a binary file
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            
            filename = os.path.split(att_cert[name])[-1]

            # Encode to base64
            encoders.encode_base64(part)
            # Add header 
            part.add_header(
                "Content-Disposition", f"attachment; filename= {filename}",
            )
            # Add attachment to your message and convert it to string
            msg.attach(part)
            text = msg.as_string()
         
        recipient_email = att_email[name]
        
        # server.sendmail( sender_email, recipient_email,text )
        
        msg = MIMEMultipart()
        
        headerformatter("", 86, [1,0])
        #print(msg)
        print ("\nSuccessfully sent email:  {:}\t\t{:}".format(att_email[name], filename))
        headerformatter("", 86, [1,0])

# server.quit()


#this code can be use for email vr
# while True:
    
#     sender_email = requestInput("Please enter your " + str(server_name[server_choice]) + " email address")
    
#     if validateEmail(sender_email) == True:
#         headerformatter("", 86, [1,0])
#         break
#     #print(sender_email)
#     headerformatter("", 86, [1,0])
    