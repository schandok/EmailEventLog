'''
    This module emails Windows event logs since the last saved time stamp. If no time stamp exists (1st time), last 10 events will be sent.
    This module can be attached as a task in Windows Event Viewer to be executed when certain event is logged.
    Configure the params below to macth the event that triggered the module, such as logger and severity level (in cmdSinceLastTimeStamp).
    All event details will be emailed. E.g: executing below test command on command line, will show last 10 messages in System log with Level = Critical, Error or Warning. This would be the eact content of the email.

    Test command: wevtutil qe "System" "/q:*[System [(Level<4)]]" /f:text /rd:true /c:10"

    Developed on Windows 8.1, using Python 2.7.8
'''

import os
import re
import time
import smtplib
import datetime

# Configure params
# In case using gmail for email, create a new App specific password to avoid entering your main password. https://support.google.com/mail/answer/1173270?hl=en
TIMESTAMP_DIR = 'C:\\EmailEventLog'
LOGGER = 'System'     # Get list of loggers by executing 'wevtutil el' on command line.
SMTP_HOST = 'smtp.gmail.com'
SMTP_PORT = 587
SEND_TO   = 'your_email@gmail.com'
USERNAME  = 'your_email@gmail.com'
PASSWORD  = 'your_passwd'

# Commands to query event log information. cmdSinceLastTimeStamp is querying events AFTER last timestamp and level (1-3) - Critical, Error, and Warning.
# To write a more custom wevtutil query, go to http://technet.microsoft.com/en-us/library/cc732848.aspx
cmdSinceLastTimeStamp = "wevtutil qe \"" + LOGGER + "\" \"/q:*[System [(Level<4) and TimeCreated[@SystemTime>'{0}']]]\" /f:text /rd:true"
cmdLast10 = "wevtutil qe \"" + LOGGER + "\" \"/q:*[System [(Level<4)]]\" /f:text /rd:true /c:10"

class EmailEventLog:
    def __init__(self):
        self._TIMESTAMP_FILE = TIMESTAMP_DIR + '\\' + 'timeStamp.txt'
    
    def __exit__(self):
        pass
    
    def handleEvent(self):
        lastEventTimeStamp = self._getLastEventTimeStamp()
        
        cmd = None
        # If no timestamp, query last 10.
        if None == lastEventTimeStamp:
            cmd = cmdLast10
        else:
            cmd = cmdSinceLastTimeStamp.format(lastEventTimeStamp)
        
        # Get events information from cmd line
        evtLog = os.popen(cmd).read()
        
        timeStamp = None
        # Find the latest time stamp
        for line in evtLog.split("\n"):
            if 0 == line.strip().find("Date:"):
                timeStamp = line.split(":", 1)[1].strip()
                break
        
        # If no timestamp, consider it as a false alarm
        if None == timeStamp:
            return
        
        # Send email
        self._sendEmail(evtLog)
 
        # Save new time stamp
        # NOTE: If send mail throws an exception, implying an error in sending email; timestamp won't be updated (intentionally).
        self._updateTimeStamp(timeStamp)
    
    def _getLastEventTimeStamp(self):
        ''' wevtutil expects the time stamp w/ UTC offset. Below logic adds that offset.
        '''
        try:
            timeStamp = open(self._TIMESTAMP_FILE, 'r').read().strip()
            # Ensure timestamp is in valid format, e.g: '2014-08-28T22:29:24.000'
            if None == re.match('^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}$', timeStamp):
                return
            
            timeStamp = timeStamp.split('T')
            date  = timeStamp[0].split('-')
            clock = timeStamp[1].replace('.', ':').split(':')
            offset = time.timezone if time.localtime().tm_isdst == 0 else time.altzone
            utcTimeStamp = datetime.datetime(int(date[0]), int(date[1]), int(date[2]), int(clock[0]), int(clock[1]), int(clock[2]), int(clock[3])*1000) + datetime.timedelta(0, offset)
            
            return utcTimeStamp.isoformat() if int(clock[3]) == 0 else utcTimeStamp.isoformat()[:-3]
        except IOError:
            pass
    
    def _updateTimeStamp(self, timeStamp):
        open(self._TIMESTAMP_FILE, 'w+').write(timeStamp)
    
    def _sendEmail(self, msg):
        smtpServer = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        smtpServer.ehlo()
        smtpServer.starttls()
        smtpServer.ehlo
        smtpServer.login(USERNAME, PASSWORD)
        header = 'To: {0}\nFrom: {1}\nSubject: {2} \n\n'.format(SEND_TO, USERNAME, 'Mail from EmailEventLog')
        smtpServer.sendmail(USERNAME, SEND_TO, header + msg)
        smtpServer.close()

if '__main__' == __name__:
    EmailEventLog().handleEvent()
