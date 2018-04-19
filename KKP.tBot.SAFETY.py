# KKP.tBot.py   (Version 0.2)
# Created by Michael Long in Python (Version 2.7)
#
# Intended to increase attendance and communication between the chapter its 
# brothers that have twitter. Tweets the events for the week, given a document 
# name of meeting minutes. Events are identifiec with a ">".
#      (Can be easily installed using command line 'pip install <LIBRARYNAME>')
#
#                        USED LIBRARIES:
# tweepy is a library to simplify the twitter access and API
# python-docx can process microsoft word documents with a .docx extension
# sys is a library that contains multiple system functions
# datetime is a library to aid with date representation of events
import tweepy
from docx import Document
import sys
import datetime

#                       Email Libraries
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime import base as MIMEBase
from email import encoders as Encoders

#Twitter API Information found on Twitter App and on Developer page
CONSUMER_KEY = #PRIVATE
CONSUMER_SECRET = #PRIVATE
ACCESS_KEY = #PRIVATE
ACCESS_SECRET = #PRIVATE

#Setting tweepy Authentifiers for the API access  
authent = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
authent.set_access_token(ACCESS_KEY, ACCESS_SECRET)
api = tweepy.API(authent)

#Console Introduction and Instructions for User
print "\n\n\nWelcome to KKP.tBot.py, written by Michael Long"
print "\nFile input is case sensitive and requires the document "
print "extension identifier (.docx)"

#Ensures file input to program contains a .docx extension, Also checks for 
#file's existance/ availability.
sentinel = 0
while(sentinel != 500):
        fileName = raw_input("Please Enter Name of File: ")
        exitCase = fileName.lower()
        if(exitCase == "exit"):
            print "\n\n Thank you for using KKP.tBot.py written by "
            print "Michael Long in Python 2.7"
            raw_input("\n\nPress any key to terminate program.")
            sys.exit(0)
        if (fileName.__contains__(".docx")):
            sentinel = 500
        else:
            print "File extension not valid, .docx is required!\n"
            print "Enter \"exit\" to end program."

#Document Defining and Handles file not found exception!
try:
    document = Document(fileName)
    docStr = []
    i = 0
    for p in document.paragraphs:
        i += 1
        docStr.append(p.text)
    print "\nMaximum Index of Events: " + str(i)
except OSError:
    print "\nFile name not found.\n\n Terminating program."
    raw_input("\n\nPress any key to terminate program.")
    sys.exit(0)

#                              INPUT PARSER
# Parses document for paragraphs containing "EVENT" keyword, and then splits 
# the rest of the paragraph using "*" character as a separator of input.
k = 0
p = 0
eventArray = []
parseStr = []
body = "Brother:\n\n_______________________________"
while k in range(0,i):
    rawStr = docStr[k]
    if(rawStr.__contains__("EVENT")):
        if(p == 0):
            currentDate = datetime.date.today().strftime("%B %d, %Y")
            weekStr = "Events for the week of " + currentDate
            body += weekStr + "_______________________________\n"
            print weekStr
            #api.update_status(weekStr)                 #Prints week of events
        parseStr = rawStr.split("*")
        nm = parseStr[1]
        dt = parseStr[2]
        lc = parseStr[3]
        tm = parseStr[4]
        ds = parseStr[5]
        formStr = nm + " - " + dt + " - " + lc + " - " + tm + " - " + ds
        body += "\n" + formStr
        #api.update_status(formStr)      #Creates individual strings of events
        print formStr.encode('utf-8')
        p += 1
    k += 1

#                       EMAILING MEETING MINUTES!
# The portions below and above create the text in the email. Gives a rundown 
# of upcoming events that week and describes the purpose of the email.
body += "\n\nAttached to this email are the meeting minutes." 
body += "\n\nAEA,\n\tMike Long \n\n\nThis email was automatically generated "
body += " and sent with a python script, https://github.com/mlong4UNH/NO-Bot"
body += "\nPlease respond to this email if any errors have occurred."
print body

#Begin connection to email server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.ehlo()
server.starttls()

#Entering gmail account and password and sigining in
x = 0
while(x != 500):
    senderEmail = raw_input("Please Enter gmail to send from: ")
    testStr1 = senderEmail.lower()
    if(testStr1 == 'quit'):
        print "\n\n Thank you for using KKP.tBot.py written by "
        print "Michael Long in Python 2.7"
        raw_input("\n\nPress any key to terminate program.")
        sys.exit(0)
    senderPass = raw_input("\nPlease Enter account's password: ")
    testStr2 = senderPass.lower()
    if(testStr2 == 'quit'):
        print "\n\n Thank you for using KKP.tBot.py written by "
        print "Michael Long in Python 2.7"
        raw_input("\n\nPress any key to terminate program.")
        sys.exit(0)
    try:
        server.login(senderEmail,senderPass)
        x = 500
    except:
        print "An Error Occured, please re-enter information or type /'quit/' to end program"

print "Successful Connection"

#TEMPORARY
recipEmail = #PRIVATE

#Assemble completed email and prepare to send
message = MIMEMultipart()
subjectLine = 'Meeting Minutes for %s' % (currentDate)
message['Subject'] = subjectLine
message['From'] = senderEmail
message['To'] = recipEmail
message.attach(MIMEText(body))

#Make a part to send an email to all brothers here!
server.sendmail(senderEmail, recipEmail, message.as_string()) #Send assembled email to designated recipient.
#END RECIPIENT BLOCK

#Finish sending emails here
server.close()

raw_input("\n\nPress any key to terminate program.")
