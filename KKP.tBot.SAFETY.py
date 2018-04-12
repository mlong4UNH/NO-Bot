# KKP.tBot.py
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
from docx import Document
import tweepy
import sys
import datetime

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
print "Welcome to KKP.tBot.py, written by Michael Long"
print "\nFile input is case sensitive and requires the document "
print "extension identifier (.docx)"

#Ensures file input to program contains a .docx extension, Also checks for 
#file's existance/ availability.
sentinel = 0
while(sentinel != 500):
    fileName = raw_input("Please Enter Name of File: ")
    if(fileName == "exit"):
        print "\n\nTerminating program.\n Thank you for using KKP.tBot.py " 
        print "written by Michael Long in Python 2.7"
        sys.exit(0)
    if (fileName.__contains__(".docx")):
        print "File extension is valid, now searching for file"
        sentinel = 500
    else:
        print "File extension not valid, .docx is required!\n"
        print "Enter \"exit\" to end program."
    
#Document Defining
document = Document(fileName)
docStr = []
i = 0
for p in document.paragraphs:
    i = i + 1
    docStr.append(p.text)
print "\nMaximum Index of Events: " + str(i)

#Parses document for paragraphs containing "EVENT" keyword, and then splits the
#rest of the paragraph using "-" character as a separator of input.
#
#                                  EXAMPLE INPUT:
#EVENT - Pie a Psi - April 13, 2018 - 10:00am to 12:00pm - Pie a brother!
k = 0
p = 0
rawStr = ""
while k in range(0,i):
    rawStr = docStr[k]
    if(rawStr.__contains__("EVENT")):
        if(p == 0):
            currentDate = datetime.datetime.now()
            WeekStr = "Events for the week of " + currentDate
            api.update_status(WeekStr)
        parseStr = rawStr.split(" - ")
        nm = parseStr[1]
        lc = parseStr[2]
        dt = parseStr[3]
        tm = parseStr[4]
        ds = parseStr[5]
        formStr = nm + " - " + lc + " - " + dt + " - " + tm + " - " + ds
        api.update_status(formStr)
        print "Successful Tweet: " + formStr
        p = 1
    k = k + 1

#TO DO: Check for file's existance/availability prior to attempting to open it.
#TO DO: Use datetime to more accurately relay date information.
#TO DO: Put in fail safes for forbidden words