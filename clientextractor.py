'''

Note: This module uses source code provided by Google Inc.
The original oauth2.py script can be found at:
    https://github.com/google/gmail-oauth2-tools/blob/master/python/oauth2.py

'''

import imaplib
import email
import lxml
import sys
import urllib
import json
import re
import datetime
import xlsxwriter

from os import listdir

import calendar
from pandas import read_html

from optparse import OptionParser

# Gmail credentials file path
CREDENTIALS_PATH = "./creds_filled.data"

# The URL root for accessing Google Accounts.
GOOGLE_ACCOUNTS_BASE_URL = 'https://accounts.google.com'

# Hardcoded dummy redirect URI for non-web apps.
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

#SCOPE= 'https://www.googleapis.com/auth/gmail.readonly'  
SCOPE= 'https://mail.google.com/'

ENDOFHEADER= "Number\r\n"

MONTH_ABBR_NUMBERS = {v + ".": k for k,v in enumerate(calendar.month_abbr)}
MONTH_NUMBERS = {v: k for k,v in enumerate(calendar.month_name)}
MONTH_NUMBERS_INVERSE = {k: v for k,v in enumerate(calendar.month_name)}
MONTH_ABBR_NUMBERS_INVERSE = {k: v for k,v in enumerate(calendar.month_abbr)}


CURRENTYEAR = 2017

class Client:
    
    def __init__(self, ref_number, update_datetime, created_datetime, firstName, 
            lastName, email, airlines, flight_number, origin, arrival_datetime, arrival_weekday ):
        self.ref_number= ref_number
        self.update_datetime= update_datetime
        self.created_datetime= created_datetime
        self.firstName= firstName
        self.lastName= lastName
        self.email= email
        self.airlines= airlines
        self.flight_number= flight_number
        self.origin= origin
        self.arrival_datetime= arrival_datetime
        self.arrival_weekday= arrival_weekday

    
    def GetDataSetAsList():
        return [self.firstName, self.lastName, self.flight_number, self.arrival_weekday,  self.getArrivalDateAsString(),
                                self.getArrivalTimeAsString, "TODO", "TODO", "TODO", "TODO"]


    def setDateTimeLastUpdated(self, year, month, day, hour, minute):
        self.dateTimeUpdated = datetime.datetime(year, month, day, hour, minute)

    def setDateTimeCreated(self, year, month, day, hour, minute):
        self.dateTimeCreated = datetime.datetime(year, month, day, hour, minute)

    def setReferenceNumber(self, refNumber):
        self.referenceNumber = refNumber

    def setFirstName(self, firstName):
        self.firstName = firstName

    def setLastName(self, lastName):
        self.lastName = lastName

    def getArrivalDateAsString(self):
        return self.arrival_datetime.day + " " + MONTH_ABBR_NUMBERS_INVERSE[self.arrival_datetime.day] + " " + self.arrival_datetime.year 

    def getArrivalTimeAsString(self):
        return self.arrival_datetime.hour + ":" + self.arrival_datetime.minute

class ClientExtractor:
    def __init__(self):
        self.auth_string= ""
        self.clientSet = []

    def ExecuteSequence(self):
        self.InitializeCredentials()
        self.GetRawClientList()
        self.ConvertListToClients()
        self.WriteSpreadsheet()
    
    def InitializeCredentials(self):
        credentials = MiscTools.GetGmailCreds(CREDENTIALS_PATH)
        self.username = credentials['USERNAME']
        self.client_id = credentials['CLIENTID']
        self.client_token = credentials['CLIENTTOKEN']

    def GetRawClientList(self):
        if self.auth_string ==  "":
            print 'To authorize token, visit this url and follow the directions:'
            print '  %s' % OAuth2Tools.GeneratePermissionUrl(self.client_id, SCOPE)
            authorization_code = raw_input('Enter verification code: ')
            auth_tokens= OAuth2Tools.AuthorizeTokens(self.client_id, self.client_token, authorization_code)
            self.auth_string= OAuth2Tools.GenerateOAuth2String(self.username, auth_tokens['access_token'], base64_encode=False)
        latestRawEmail= EmailTools.GetLatestEmail(self.username,  self.auth_string)
        latestEmail= EmailTools.ConvertRawToEmailMessage(latestRawEmail)
        emailData= EmailTools.ConvertEmailMessageToData(latestEmail, 0)
        self.dataList = DataTools.BreakDataStringToDataList(emailData)

    def GetFakeRawClientList(self):
        with open("./data/EmailData2.txt",'r') as emailData:
            self.dataList = emailData.read()

    def ConvertListToClients(self):
        
        month = int(raw_input('Enter Month #: '))
        day = int(raw_input('Enter Day #: '))

        dateToGet = datetime.date(CURRENTYEAR, month, day)
        
        self.dateFound = dateToGet

        self.clientSet = DataTools.HtmlStringToClientList(self.dataList, dateToGet)

    def WriteSpreadsheet(self):
        updateNumber = 1
        filePrefix = 'students_' + self.dateFound.day + MONTH_NUMBERS_INVERSE[self.dateFound.month] + '_update' 
        directoryFileList = listdir("./spreadsheets")
        for eachFile in directoryFileList:
            if filePrefix in eachFile:
                updateNumber++
        workbook = xlsxwriter.Workbook( filePrefix  + updateNumber  + '.xlsx')
        worksheet = workbook.add_worksheet()
        headers = ["First name", "Family name", "Airline Flight No.", "Arrival day Arrival Date",
                "Arrival time (est.)", "Extra passengers", "Drop-off (University Residence)", 
                "Drop-off Address (other)", "Suburb"]
        numberClients = len(clientSet)
        for col,eachHeader in enumerate(headers):
            worksheet.write(0, col, eachHeader)
        for row, eachClient in enumerate(self.clientSet):
            for col, eachData in enumerate(eachClient.GetDataSetAsList())
                worksheet.write(row + 1, col, eachData)
        workbook.close()


    #check if flightnumbers match flights
    #highlight updates
class DataTools:

    @staticmethod
    def HtmlStringToClientList(html_string, date):
            clientList = []
            htmlParsed = re.split("<tr>|</tr>", html_string)
            betterParsed = htmlParsed[18:-3:2]
            for client in betterParsed:
                currentClient = read_html("<table>" + "<tr>" + client + "</tr>" + "</table>")[0].values.tolist()[0]
                pickUpDate = currentClient[16]
                ###print pickUpDate
                if pickUpDate == u'\xc2':
                    continue
                pickUpDateParsed = pickUpDate.split()
                monthName = pickUpDateParsed[1]
                if "." in monthName:
                    pickUpDateObject = datetime.date(int(pickUpDateParsed[2]), MONTH_ABBR_NUMBERS[monthName], int(pickUpDateParsed[0]))
                else:
                    pickUpDateObject = datetime.date(int(pickUpDateParsed[2]), MONTH_NUMBERS[monthName], int(pickUpDateParsed[0]))
                if pickUpDateObject == date:
                    pickUpTime = MiscTools.TimeStringToTimeObject(currentClient[17])
                    pickUpDateTime = datetime.datetime(pickUpDateObject.year, pickUpDateObject.month, pickUpDateObject.day, pickUpTime.hour, pickUpTime.minute)
                    newClient = Client(currentClient[1], MiscTools.DateTimeStringToDateTimeObjects(currentClient[2]),
                            MiscTools.DateTimeStringToDateTimeObjects(currentClient[3]), currentClient[4], currentClient[5],
                            currentClient[6], currentClient[12], currentClient[13], currentClient[14], pickUpDateTime, currentClient[15])
                    clientList.append(newClient)
            return clientList


    @staticmethod
    def BreakDataStringToDataList(dataString):
        dataList = re.split('[0-9][0-9][0-9][0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9][0-9][0-9]', dataString)
        return dataList

    @staticmethod
    def SplitFirstWordOffString(string_to_split):
        string_to_split = string_to_split.lstrip()
        first_word = string_to_split.split(" ")[0]
        new_string = string_to_split.replace(first_word,"")
        new_string = new_string.lstrip()
        return first_word, new_string

class MiscTools:
    
    @staticmethod
    def DateTimeStringToDateTimeObjects(dateTimeString):
        #print dateTimeString
        date, time = dateTimeString.split('  ')
        day,month,year = date.split('/')
        timeNumber, amPm  = time.split(' ')
        hour, minute = timeNumber.split('.')
        if amPm == u"PM":
            hour = int(hour) + 12
            if hour == 24:
                hour = 0
        return datetime.datetime(int(year), int(month), int(day), int(hour), int(minute))

    @staticmethod
    def TimeStringToTimeObject(timeString):
        hour = int(timeString[0:2])
        if timeString[3:5] == '':
            minute = 0
        else:
            minute = int(timeString[3:5])
        if 'PM' in timeString :
            hour = hour + 12
        return datetime.time(hour, minute)

    @staticmethod
    def DatesAreCloseEnough(date1, date2, distanceInDays):
        pass

    @staticmethod
    def GetGmailCreds(path_to_data_file):
        credentials = {}
        with open(path_to_data_file, 'r') as credsFile:
            for line in credsFile:
                (key, val) = line.split('=')
                key = key.replace(" ","")
                val = val.replace("\n","")
                val = val.replace(" ","")
                credentials[key] = val
        return credentials

class EmailTools:

    @staticmethod
    def GetLatestEmail(EMAILUSER,  auth_string):
        imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
        imap_conn.debug = 4
        imap_conn.authenticate('XOAUTH2', lambda x: auth_string)
        imap_conn.select('INBOX')
        result, data = imap_conn.uid('search', None, "ALL") # search and return uids instead
        latest_email_uid = data[0].split()[-1]
        result, data = imap_conn.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = data[0][1]
        return raw_email

    @staticmethod
    def ConvertRawToEmailMessage(raw_email):
        return email.message_from_string(raw_email)

#TODO:save text files of both raw emails to avoid data cap
    @staticmethod
    def ConvertEmailMessageToData(email_message, payload_index):
        emailPayload= email_message.get_payload(payload_index)
        dataDecodeable= emailPayload.get_payload(decode= True)
        dataDecoded= dataDecodeable.decode('utf-8')
        return dataDecoded
        #startDataIndex= dataDecoded.find(ENDOFHEADER)
        #return dataDecoded[(startDataIndex + len(ENDOFHEADER)):]



class OAuth2Tools:

    @staticmethod
    def AccountsUrl(command):
        """Generates the Google Accounts URL.

        Args:
            command: The command to execute.

        Returns:
            A URL for the given command.
        """
        return '%s/%s' % (GOOGLE_ACCOUNTS_BASE_URL, command)


    @staticmethod
    def UrlEscape(text):
        # See OAUTH 5.1 for a definition of which characters need to be escaped.
        return urllib.quote(text, safe='~-._')


    @staticmethod
    def UrlUnescape(text):
        # See OAUTH 5.1 for a definition of which characters need to be escaped.
        return urllib.unquote(text)


    @staticmethod
    def FormatUrlParams(params):
        """Formats parameters into a URL query string.

        Args:
            params: A key-value map.

        Returns:
            A URL query string version of the given parameters.
        """
        param_fragments = []
        for param in sorted(params.iteritems(), key=lambda x: x[0]):
            param_fragments.append('%s=%s' % (param[0], OAuth2Tools.UrlEscape(param[1])))
        return '&'.join(param_fragments)


    @staticmethod
    def GeneratePermissionUrl(client_id, scope='https://mail.google.com/'):
        """Generates the URL for authorizing access.

        This uses the "OAuth2 for Installed Applications" flow described at
        https://developers.google.com/accounts/docs/OAuth2InstalledApp

        Args:
            client_id: Client ID obtained by registering your app.
            scope: scope for access token, e.g. 'https://mail.google.com'
        Returns:
            A URL that the user should visit in their browser.
        """
        params = {}
        params['client_id'] = client_id
        params['redirect_uri'] = REDIRECT_URI
        params['scope'] = scope
        params['response_type'] = 'code'
        return '%s?%s' % (OAuth2Tools.AccountsUrl('o/oauth2/auth'), OAuth2Tools.FormatUrlParams(params))


    @staticmethod
    def AuthorizeTokens(client_id, client_secret, authorization_code):
        """Obtains OAuth access token and refresh token.

        This uses the application portion of the "OAuth2 for Installed Applications"
        flow at https://developers.google.com/accounts/docs/OAuth2InstalledApp#handlingtheresponse

        Args:
            client_id: Client ID obtained by registering your app.
            client_secret: Client secret obtained by registering your app.
            authorization_code: code generated by Google Accounts after user grants
            permission.
        Returns:
            The decoded response from the Google Accounts server, as a dict. Expected
            fields include 'access_token', 'expires_in', and 'refresh_token'.
        """
        params = {}
        params['client_id'] = client_id
        params['client_secret'] = client_secret
        params['code'] = authorization_code
        params['redirect_uri'] = REDIRECT_URI
        params['grant_type'] = 'authorization_code'
        request_url = OAuth2Tools.AccountsUrl('o/oauth2/token')

        response = urllib.urlopen(request_url, urllib.urlencode(params)).read()
        return json.loads(response)


    @staticmethod
    def RefreshToken(client_id, client_secret, refresh_token):
        """Obtains a new token given a refresh token.

        See https://developers.google.com/accounts/docs/OAuth2InstalledApp#refresh

        Args:
            client_id: Client ID obtained by registering your app.
            client_secret: Client secret obtained by registering your app.
            refresh_token: A previously-obtained refresh token.
        Returns:
            The decoded response from the Google Accounts server, as a dict. Expected
            fields include 'access_token', 'expires_in', and 'refresh_token'.
        """
        params = {}
        params['client_id'] = client_id
        """Generates an IMAP OAuth2 authentication string.

        See https://developers.google.com/google-apps/gmail/oauth2_overview

        Args:
            username: the username (email address) of the account to authenticate
            access_token: An OAuth2 access token.
            base64_encode: Whether to base64-encode the output.

        Returns:
            The SASL argument for the OAuth2 mechanism.
        """
        auth_string = 'user=%s\1auth=Bearer %s\1\1' % (username, access_token)
        if base64_encode:
            auth_string = base64.b64encode(auth_string)
        return auth_string


    @staticmethod
    def TestImapAuthentication(user, auth_string):
        """Authenticates to IMAP with the given auth_string.

        Prints a debug trace of the attempted IMAP connection.

        Args:
            user: The Gmail username (full email address)
            auth_string: A valid OAuth2 string, as returned by GenerateOAuth2String.
                Must not be base64-encoded, since imaplib does its own base64-encoding.
        """
        print
        imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
        imap_conn.debug = 4
        imap_conn.authenticate('XOAUTH2', lambda x: auth_string)
        imap_conn.select('INBOX')


    @staticmethod
    def TestSmtpAuthentication(user, auth_string):
        """Authenticates to SMTP with the given auth_string.

        Args:
            user: The Gmail username (full email address)
            auth_string: A valid OAuth2 string, not base64-encoded, as returned by
                GenerateOAuth2String.
        """
        print
        smtp_conn = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_conn.set_debuglevel(True)
