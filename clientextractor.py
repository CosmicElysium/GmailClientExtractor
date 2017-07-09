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
import time

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

class Client
    
    def setDateLastUpdated:
        pass

    def setDateCreated:
        pass



class ClientExtractor:
    def __init__(self):
        self.auth_string= ""
        self.clientSet = []

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

    def ConvertListToClients
        for eachStudent in self.dataList:
            entryDates = re.search('([0-9][0-9]\/[0-9][0-9]\/[0-9][0-9][0-9][0-9]).*?([0-9][0-9]\/[0-9][0-9]\/[0-9][0-9][0-9][0-9])', eachStudent)
            dateUpdated = entryDates.group(1)
            dateCreated = entryDates.group(2)
            


class DataTools:

    @staticmethod
    def BreakDataStringToDataList(dataString):
        dataList = re.split('[0-9][0-9][0-9][0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9][0-9][0-9]', dataString)
        return dataList

class MiscTools:
    
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
        params['client_secret'] = client_secret
        params['refresh_token'] = refresh_token
        params['grant_type'] = 'refresh_token'
        request_url = OAuth2Tools.AccountsUrl('o/oauth2/token')

        response = urllib.urlopen(request_url, urllib.urlencode(params)).read()
        return json.loads(response)


    @staticmethod
    def GenerateOAuth2String(username, access_token, base64_encode=True):
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
