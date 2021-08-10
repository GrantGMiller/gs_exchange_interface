'''
All datetimes that are passed to/from this module are in the system local time.

'''
import datetime
import json
import re
from base64 import b64decode
import time
from gs_service_accounts import _ServiceAccountBase

try:
    from extronlib.system import ProgramLog
    import gs_requests as requests
except Exception as e:
    print(str(e))
    import requests

from gs_calendar_base import (
    _BaseCalendar,
    CalendarItem,
    ConvertDatetimeToTimeString,
    ConvertTimeStringToDatetime
)
import gs_oauth_tools

TZ_NAME = time.tzname[0]
if TZ_NAME == 'EST':
    TZ_NAME = 'Eastern Standard Time'
elif TZ_NAME == 'PST':
    TZ_NAME = 'Pacific Standard Time'
elif TZ_NAME == 'CST':
    TZ_NAME = 'Central Standard Time'
########################################

RE_CAL_ITEM = re.compile('<t:CalendarItem>[\w\W]*?<\/t:CalendarItem>')
RE_ITEM_ID = re.compile(
    '<t:ItemId Id="(.*?)" ChangeKey="(.*?)"/>'
)  # group(1) = itemID, group(2) = changeKey #within a CalendarItem
RE_SUBJECT = re.compile('<t:Subject>(.*?)</t:Subject>')  # within a CalendarItem
RE_HAS_ATTACHMENTS = re.compile('<t:HasAttachments>(.{4,5})</t:HasAttachments>')  # within a CalendarItem
RE_ORGANIZER = re.compile(
    '<t:Organizer>.*<t:Name>(.*?)</t:Name>.*</t:Organizer>'
)  # group(1)=Name #within a CalendarItem
RE_START_TIME = re.compile('<t:Start>(.*?)</t:Start>')  # group(1) = start time string #within a CalendarItem
RE_END_TIME = re.compile('<t:End>(.*?)</t:End>')  # group(1) = end time string #within a CalendarItem
RE_HTML_BODY = re.compile('<t:Body BodyType="HTML">([\w\W]*)</t:Body>', re.IGNORECASE)

RE_EMAIL_ADDRESS = re.compile(
    "[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
)

RE_ERROR_CLASS = re.compile('ResponseClass="Error"', re.IGNORECASE)
RE_ERROR_MESSAGE = re.compile('<m:MessageText>([\w\W]*)</m:MessageText>')

RE_ATTENDEE = re.compile('<t:Attendee>([\w\W]*)</t:Attendee>')

RE_CAL_FOLDER = re.compile('<t:CalendarFolder>([\w\W]*?)<\/t:CalendarFolder>')
RE_FOLDER_ID = re.compile('<t:FolderId Id="(.*?)" ChangeKey="(.*?)"/>')
RE_PARENT_FOLDER_ID = re.compile('<t:ParentFolderId Id="(.*?)" ChangeKey="(.*?)"/>')


class EWS(_BaseCalendar):
    def __init__(
            self,
            username=None,
            password=None,
            impersonation=None,
            myTimezoneName=None,
            serverURL=None,
            authType='Basic',  # also accept "NTLM" and "Oauth"
            oauthCallback=None,  # callable, takes no args, returns Oauth token
            apiVersion='Exchange2007_SP1',  # TLS uses "Exchange2007_SP1"
            verifyCerts=True,
            debug=False,
            persistentStorage=None,
    ):
        self.username = username
        self.password = password
        self.impersonation = impersonation
        self._serverURL = serverURL
        self._authType = authType
        self._oauthCallback = oauthCallback
        self._apiVersion = apiVersion
        self._verifyCerts = verifyCerts
        self._debug = debug

        #################################################################
        thisMachineTimezoneName = time.tzname[0]
        if thisMachineTimezoneName == 'EST':
            thisMachineTimezoneName = 'Eastern Standard Time'
        elif thisMachineTimezoneName == 'PST':
            thisMachineTimezoneName = 'Pacific Standard Time'
        elif thisMachineTimezoneName == 'CST':
            thisMachineTimezoneName = 'Central Standard Time'

        self._myTimezoneName = myTimezoneName or thisMachineTimezoneName
        if self._debug: print('myTimezoneName=', self._myTimezoneName)

        self._session = requests.session()

        self._session.headers['Content-Type'] = 'text/xml'

        if callable(oauthCallback) or authType == 'Oauth':
            self._authType = authType = 'Oauth'
        elif authType == 'Basic':
            self._session.auth = requests.auth.HTTPBasicAuth(self.username, self.password)
        else:
            raise TypeError('Unknown Authorization Type')
        self._useImpersonationIfAvailable = True
        self._useDistinguishedFolderMailbox = False

        self._batchQueue = []
        self._folderIDs = {
            # str(email): str(folderID)
        }
        self._otherEWSInstances = {
            # str(email), EWS()
        }
        super().__init__(
            name=impersonation or username,
            persistentStorage=persistentStorage,
            debug=debug
        )

    def print(self, *a, **k):
        if self._debug:
            print(*a, **k)

    def __str__(self):
        if self._oauthCallback:
            return '<EWS: state={}, impersonation={}, auth={}, oauthCallback={}{}>'.format(
                self._connectionStatus,
                self.impersonation,
                self._authType,
                self._oauthCallback,
                ', id={}'.format(id(self)) if self._debug else '',
            )
        else:
            return '<EWS: state={}, username={}, impersonation={}, auth={}{}>'.format(
                self._connectionStatus,
                self.username,
                self.impersonation,
                self._authType,
                ', id={}'.format(id(self)) if self._debug else '',
            )

    @property
    def Impersonation(self):
        return self.impersonation

    @Impersonation.setter
    def Impersonation(self, newImpersonation):
        self.impersonation = newImpersonation

    def DoRequest(self, soapBody, truncatePrint=False, tryNumber=0):
        # API_VERSION = 'Exchange2013'
        # API_VERSION = 'Exchange2007_SP1'

        if self.impersonation and self._useImpersonationIfAvailable:
            # Note: Don't add a namespace to the <ExchangeImpersonation> and <ConnectingSID> tags
            # This will cause a "You don't have permission to impersonate this account" error.
            # Don't ask my why.
            # UPDATE: removing the namespace makes this work for licensed accounts, but not for service accounts with impersonation, so now i really dont understand whats going on

            # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-identify-the-account-to-impersonate
            soapHeader = '''
                <t:RequestServerVersion Version="{apiVersion}" />
                <t:ExchangeImpersonation>
                    <t:ConnectingSID>
                        <t:PrimarySmtpAddress>{impersonation}</t:PrimarySmtpAddress> <!-- Needs to be in a single line -->
                    </t:ConnectingSID>
                </t:ExchangeImpersonation>
            '''.format(
                apiVersion=self._apiVersion,
                impersonation=self.impersonation,
            )
        else:
            soapHeader = '<t:RequestServerVersion Version="{apiVersion}" />'.format(apiVersion=self._apiVersion)

        soapEnvelopeOpenTag = '''
            <soap:Envelope 
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
            >'''

        xml = '''<?xml version="1.0" encoding="utf-8"?>
                    {soapEnvelopeOpenTag}
                        <soap:Header>
                            {soapHeader}
                        </soap:Header>
                        <soap:Body>
                            {soapBody}
                        </soap:Body>
                    </soap:Envelope>
        '''.format(
            soapEnvelopeOpenTag=soapEnvelopeOpenTag,
            soapHeader=soapHeader,
            soapBody=soapBody,
        )

        self.print('xml=', xml)

        if self._serverURL:
            url = self._serverURL + '/EWS/exchange.asmx'
        else:
            url = 'https://outlook.office365.com/EWS/exchange.asmx'

        if self._authType == 'Oauth':
            self._session.headers['Authorization'] = 'Bearer {token}'.format(token=self._oauthCallback())

        if self._debug:
            for k, v in self._session.headers.items():
                if 'auth' in k.lower():
                    v = v[:15] + '...'
                self.print('header', k, v)

        resp = self._session.request(
            method='POST',
            url=url,
            data=xml,
            verify=self._verifyCerts,
        )
        if self._debug:
            print('resp.status_code=', resp.status_code)
            print('resp.reason=', resp.reason)
            if truncatePrint:
                print('resp.text=', resp.text[:1024])
            else:
                print('resp.text=', resp.text)

        if resp.ok and RE_ERROR_CLASS.search(resp.text) is None:
            self._NewConnectionStatus('Connected')
        else:
            for match in RE_ERROR_MESSAGE.finditer(resp.text):
                if self._debug: print('Error Message:', match.group(1))
            self._NewConnectionStatus('Disconnected')

            if 'The account does not have permission to impersonate the requested user.' in resp.text or not resp.ok or 'ErrorImpersonateUserDenied' in resp.text:
                if self._useImpersonationIfAvailable is True:
                    if self._debug: print('Switching impersonation mode')

                    self._useImpersonationIfAvailable = not self._useImpersonationIfAvailable
                    self._useDistinguishedFolderMailbox = not self._useDistinguishedFolderMailbox

                    if self._debug: print('self._useImpersonationIfAvailable=', self._useImpersonationIfAvailable)
                    if self._debug: print('self._useDistinguishedFolderMailbox=', self._useDistinguishedFolderMailbox)
                if tryNumber <= 3:
                    self.print('Trying again')
                    return self.DoRequest(
                        soapBody=soapBody,
                        truncatePrint=truncatePrint,
                        tryNumber=tryNumber + 1,

                    )
        return resp

    def UpdateCalendar(self, calendar=None, startDT=None, endDT=None):
        self.print('UpdateCalendar(', calendar, startDT, endDT)

        startDT = startDT or datetime.datetime.now() - datetime.timedelta(days=1)
        startDT = startDT.replace(second=0, microsecond=0)

        endDT = endDT or datetime.datetime.now() + datetime.timedelta(days=7)

        startTimestring = ConvertDatetimeToTimeString(startDT)
        endTimestring = ConvertDatetimeToTimeString(endDT)

        calendar = calendar or self.impersonation or self.username
        self.GetFolderInfo([calendar])

        if self._useDistinguishedFolderMailbox:
            parentFolder = '''
                <t:DistinguishedFolderId Id="calendar">
                    <t:Mailbox>
                        <t:EmailAddress>{impersonation}</t:EmailAddress>
                    </t:Mailbox>
                </t:DistinguishedFolderId>
            '''.format(
                impersonation=self.impersonation
            )
        else:
            parentFolder = '''
                <t:DistinguishedFolderId Id="calendar"/>
            '''
        # https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/fielduri
        # note: attendees must be found using <GetItem>
        soapBody = ''' 
            <m:FindItem Traversal="Shallow">
                <m:ItemShape>
                    <t:BaseShape>IdOnly</t:BaseShape>
                    <t:AdditionalProperties>
                        <t:FieldURI FieldURI="item:Subject" />
                        <t:FieldURI FieldURI="calendar:Start" />
                        <t:FieldURI FieldURI="calendar:End" />
                        <t:FieldURI FieldURI="item:Body" />
                        <t:FieldURI FieldURI="calendar:Organizer" />
                        <t:FieldURI FieldURI="item:HasAttachments" />
                        <t:FieldURI FieldURI="item:Size" />
                        <t:FieldURI FieldURI="item:Sensitivity" />
                        <t:FieldURI FieldURI="calendar:IsOnlineMeeting" />
                    </t:AdditionalProperties>
                </m:ItemShape>
                <m:CalendarView 
                    MaxEntriesReturned="100" 
                    StartDate="{startTimestring}" 
                    EndDate="{endTimestring}" 
                    />
                <m:ParentFolderIds>
                     {parentFolder}
                </m:ParentFolderIds>
            </m:FindItem>
        '''.format(
            startTimestring=startTimestring,
            endTimestring=endTimestring,
            parentFolder=parentFolder,
        )
        resp = self.DoRequest(soapBody)
        if resp.ok:
            calItems = self.CreateCalendarItemsFromResponse(resp.text)
            self.RegisterCalendarItems(
                calItems=calItems,
                startDT=startDT,
                endDT=endDT,
            )
        else:
            if 'ErrorImpersonateUserDenied' in resp.text:
                if self._debug:
                    print('Impersonation Error. Trying again with delegate access.')
                return self.UpdateCalendar(calendar, startDT, endDT)
        return resp

    def StartBatchUpdate(self, calendar=None, startDT=None, endDT=None):
        '''

        :param calendar: str > email address of the calendar
        :param startDT:
        :param endDT:
        :return:
        '''
        self._batchQueue.append({
            'calendar': calendar,
            'startDT': startDT,
            'endDT': endDT
        })

    def DoBatchUpdate(self):
        self.GetFolderInfo([
            b['calendar'] for b in self._batchQueue
        ])

        # create the batch req
        startDT = datetime.datetime.now()
        endDT = datetime.datetime.now()

        parentFolders = ''
        for batch in self._batchQueue:
            parentFolders += '''
                           <t:DistinguishedFolderId Id="calendar">
                               <t:Mailbox>
                                   <t:EmailAddress>{email}</t:EmailAddress>
                               </t:Mailbox>
                           </t:DistinguishedFolderId>
                       '''.format(
                email=batch['calendar']
            )

            startDT = min(startDT, batch['startDT'])
            endDT = max(endDT, batch['endDT'])

        soapBody = ''' 
           <m:FindItem Traversal="Shallow">
               <m:ItemShape>
                   <t:BaseShape>AllProperties</t:BaseShape>
               </m:ItemShape>
               <m:CalendarView 
                   MaxEntriesReturned="100" 
                   StartDate="{startTimestring}" 
                   EndDate="{endTimestring}" 
                   />
               <m:ParentFolderIds>
                    {parentFolders}
               </m:ParentFolderIds>
           </m:FindItem>
        '''.format(
            startTimestring=startDT.isoformat() + 'Z',
            endTimestring=endDT.isoformat() + 'Z',
            parentFolders=parentFolders,
        )
        resp = self.DoRequest(soapBody)
        self.print('418 resp=', resp.text)

        calItems = self.CreateCalendarItemsFromResponse(resp.text)

        self.RegisterCalendarItems(
            calendarNames=[b['calendar'] for b in self._batchQueue],
            calItems=calItems,
            startDT=startDT,
            endDT=endDT
        )
        self._batchQueue.clear()

    def GetEmailFromFolderID(self, fID):
        for email, folderID in self._folderIDs.items():
            if fID == folderID:
                return email
        else:
            raise KeyError(
                'Folder ID "{}" not found. self.folderIDs={}. You mightneed to call self.GetFolderInfo() first.'.format(
                    fID,
                    self._folderIDs
                ))

    def GetFolderInfo(self, emails=None):
        '''

        :param emails: list() of str(emailAddress)
        :return: dict like {
            # str(email): str(folderID)
        }
        '''
        assert type(emails) == list, '"emails" must be of type list()'
        self.print('GetFolderInfo(', emails)

        # first make sure we have all the folder IDs
        requestedFolders = []

        parentFolders = ''
        for email in emails:
            if email in self._folderIDs and self._folderIDs[email]:
                continue  # skip this one, we already have the ID

            if email not in self._folderIDs:
                requestedFolders.append(email)

                parentFolders += '''
                               <t:DistinguishedFolderId Id="calendar">
                                   <t:Mailbox>
                                       <t:EmailAddress>{email}</t:EmailAddress>
                                   </t:Mailbox>
                               </t:DistinguishedFolderId>
                           '''.format(
                    email=email
                )

        if len(parentFolders) == 0:
            self.print('We already have all the folder info:', self._folderIDs)
            return self._folderIDs

        soapBody = ''' 
                   <m:GetFolder>
                      <m:FolderShape>
                        <t:BaseShape>Default</t:BaseShape>
                      </m:FolderShape>
                      <m:FolderIds>
                        {parentFolders}
                      </m:FolderIds>
                    </m:GetFolder>
                    '''.format(
            parentFolders=parentFolders,
        )

        resp = self.DoRequest(soapBody)
        self.print('353 Getting folder Ids resp=', resp.text)

        for index, matchCalendarFolder in enumerate(RE_CAL_FOLDER.finditer(resp.text)):
            self.print('matchCalendarFolder.group(0)=', matchCalendarFolder.group(0))
            matchFolderID = RE_FOLDER_ID.search(matchCalendarFolder.group(0))
            if matchFolderID:
                self._folderIDs[requestedFolders[index]] = matchFolderID.group(1)
                self.print('found folder ID {email}={ID}'.format(
                    email=requestedFolders[index],
                    ID=matchFolderID.group(1)
                ))

        self.print('GetFolderInfo return', self._folderIDs)
        return self._folderIDs

    def CreateCalendarItemsFromResponse(self, xml):
        '''

        :param responseString:
        :return: list of calendar items
        '''
        ret = []
        for item in GetCalendarItemDataFromXML(xml):
            calItem = CalendarItem(
                startDT=item['Start'],
                endDT=item['End'],
                data=item['Data'],
                parentCalendar=self
            )
            self.print('521 calItem=', calItem)
            ret.append(calItem)

        return ret

    def CreateCalendarEvent(self, subject, startDT, endDT, body=None, attendees=None):
        self.print('CreateCalendarEvent(subject=', subject, 'body=', body, 'startDT=', startDT, 'endDT=', endDT,
                   'attendees=', attendees)

        assert isinstance(startDT, datetime.datetime)
        assert isinstance(endDT, datetime.datetime)

        startDT = startDT.replace(second=0, microsecond=0)

        startTimeString = ConvertDatetimeToTimeString(startDT)
        endTimeString = ConvertDatetimeToTimeString(endDT)

        body = body or ''
        attendees = attendees or []

        if self._useDistinguishedFolderMailbox:
            parentFolder = '''
                <t:DistinguishedFolderId Id="calendar">
                    <t:Mailbox>
                        <t:EmailAddress>{impersonation}</t:EmailAddress>
                    </t:Mailbox>
                </t:DistinguishedFolderId>
            '''.format(
                impersonation=self.impersonation,
            )
        else:
            parentFolder = '''
                <t:DistinguishedFolderId Id="calendar"/>
            '''

        soapBody = '''
            <m:CreateItem SendMeetingInvitations="SendToNone">
                <m:SavedItemFolderId>
                    {parentFolder}
                </m:SavedItemFolderId>
                <m:Items>
                    <t:CalendarItem>
                        <t:Subject>{subject}</t:Subject>
                        <t:Body BodyType="Text">{body}</t:Body>
                        <t:Start>{startTimeString}</t:Start>
                        <t:End>{endTimeString}</t:End>
                        <t:MeetingTimeZone TimeZoneName="{tzName}" />
                        {attendeesXML}
                        <t:IsOnlineMeeting>true</t:IsOnlineMeeting>
                    </t:CalendarItem>
                </m:Items>
            </m:CreateItem>
        '''.format(
            parentFolder=parentFolder,
            startTimeString=startTimeString,
            endTimeString=endTimeString,
            subject=subject,
            body=body,
            tzName=self._myTimezoneName,
            attendeesXML=GetAttendeesXML(attendees),
        )
        resp = self.DoRequest(soapBody)

    def ChangeEventTime(self, calItem, newStartDT=None, newEndDT=None):
        self.print('ChangeEventTime(', calItem, ', newStartDT=', newStartDT, ', newEndDT=', newEndDT)

        if newStartDT:
            newStartDT = newStartDT.replace(second=0, microsecond=0)

        if newEndDT:
            newEndDT = newEndDT.replace(second=0, microsecond=0)

        props = {}

        if newStartDT is not None:
            timeString = ConvertDatetimeToTimeString(
                newStartDT
            )
            props['Start'] = timeString

        if newEndDT is not None:
            timeString = ConvertDatetimeToTimeString(
                newEndDT
            )
            props['End'] = timeString

        for prop, timeString in props.items():
            soapBody = '''
                <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
                  <m:ItemChanges>
                    <t:ItemChange>
                      <t:ItemId 
                        Id="{itemID}" 
                        ChangeKey="{changeKey}" 
                        />
                      <t:Updates>
                        <t:SetItemField>
                          <t:FieldURI FieldURI="calendar:{prop}" />
                          <t:CalendarItem>
                            <t:{prop}>{timeString}</t:{prop}>
                          </t:CalendarItem>
                        </t:SetItemField>
                      </t:Updates>
                    </t:ItemChange>
                  </m:ItemChanges>
                </m:UpdateItem>
            '''.format(
                itemID=calItem.Get('ItemId'),
                changeKey=calItem.Get('ChangeKey'),
                prop=prop,
                timeString=timeString,
                tzName=self._myTimezoneName,
                # dont cuz: Error Message: An object within a change description must contain one and only one property to modify.
            )
            self.DoRequest(soapBody)

    def ChangeEventBody(self, calItem, newBody):
        self.print('ChangeEventBody(', calItem, newBody)

        soapBody = """
            <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
              <m:ItemChanges>
                <t:ItemChange>
                  <t:ItemId 
                    Id="{itemID}"
                    ChangeKey="{changeKey}" 
                    />
                  <t:Updates>
                    <t:SetItemField>
                      <t:FieldURI FieldURI="item:Body" />
                      <t:CalendarItem>
                        <t:Body BodyType="HTML">{newBody}</t:Body>
                        <t:Body BodyType="Text">{newBody}</t:Body>
                      </t:CalendarItem>
                    </t:SetItemField>
                  </t:Updates>
                </t:ItemChange>
              </m:ItemChanges>
            </m:UpdateItem>
            """.format(
            itemID=calItem.Get('ItemId'),
            changeKey=calItem.Get('ChangeKey'),
            newBody=newBody,
        )
        resp = self.DoRequest(soapBody)

    def DeleteEvent(self, calItem):
        self.print('DeleteEvent(', calItem)
        soapBody = """
                <m:DeleteItem DeleteType="HardDelete" SendMeetingCancellations="SendToNone">
                  <m:ItemIds>
                    <t:ItemId 
                        Id="{itemID}"
                        ChangeKey="{changeKey}" 
                    />
                  </m:ItemIds>
                </m:DeleteItem>
            """.format(
            itemID=calItem.Get('ItemId'),
            changeKey=calItem.Get('ChangeKey'),
        )
        resp = self.DoRequest(soapBody)

    def GetAttachments(self, calItem):
        self.print('GetAttachments(', calItem)
        # returns a list of attachment IDs

        itemId = calItem.Get('ItemId')
        regExAttachmentID = re.compile('AttachmentId Id=\"(\S+)\"')
        regExAttachmentName = re.compile('<t:Name>(.+)</t:Name>')
        xmlBody = """
                <m:GetItem>
                  <m:ItemShape>
                    <t:BaseShape>IdOnly</t:BaseShape>
                    <t:AdditionalProperties>
                      <t:FieldURI FieldURI="item:Attachments"/>
                      <t:FieldURI FieldURI="item:HasAttachments" />
                        <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                        <t:FieldURI FieldURI="calendar:OptionalAttendees" />
                    </t:AdditionalProperties>
                    
                  </m:ItemShape>
                  
                  <m:ItemIds>
                    <t:ItemId Id="{}" />
                  </m:ItemIds>
                  
                </m:GetItem>
              """.format(itemId)

        resp = self.DoRequest(xmlBody)
        self.print('540 resp=', resp.status_code)

        foundNames = regExAttachmentName.findall(resp.text)
        foundIDs = regExAttachmentID.findall(resp.text)

        ret = dict(zip(foundNames, foundIDs))

        self.print('foundNames=', foundNames)
        self.print('foundIDs=', foundIDs)
        self.print('GetAttachments ret=', ret)

        return [_Attachment(ID, name, self) for name, ID in ret.items()]

    def GetAttendees(self, calItem):
        self.print('GetAttendees(', calItem)
        # returns a list of attachment IDs

        itemId = calItem.Get('ItemId')
        xmlBody = """
                    <m:GetItem>
                      <m:ItemShape>
                        <t:BaseShape>IdOnly</t:BaseShape>
                        <t:AdditionalProperties>
                            <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                            <t:FieldURI FieldURI="calendar:OptionalAttendees" />
                        </t:AdditionalProperties>

                      </m:ItemShape>

                      <m:ItemIds>
                        <t:ItemId Id="{}" />
                      </m:ItemIds>

                    </m:GetItem>
                  """.format(itemId)

        resp = self.DoRequest(xmlBody)
        self.print('540 resp.text=', resp.text)

        ret = set()  # Use set() to prevent duplicates
        for matchEmail in RE_EMAIL_ADDRESS.finditer(resp.text):
            ret.add(matchEmail.group(0))
        self.print('GetAttendees return', ret)
        return list(ret)

    def ChangeAttendees(self, calendarItem, addList=None, removeList=None):
        self.print('ChangeAttendees(', calendarItem, 'addList=', addList, ', removeList=', removeList)
        addList = addList or []
        removeList = removeList or []

        attendees = set(self.GetAttendees(calendarItem))
        for email in addList:
            attendees.add(email)
        for email in removeList:
            attendees.remove(email)

        soapBody = """
            <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
              <m:ItemChanges>
                <t:ItemChange>
                  <t:ItemId 
                    Id="{itemID}"
                    ChangeKey="{changeKey}" 
                    />
                  <t:Updates>
                    <t:SetItemField>
                        <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                         <t:CalendarItem>
                            {attendeesXML}
                         </t:CalendarItem>
                    </t:SetItemField>
                  </t:Updates>
                </t:ItemChange>
              </m:ItemChanges>
            </m:UpdateItem>
            """.format(
            itemID=calendarItem.Get('ItemId'),
            changeKey=calendarItem.Get('ChangeKey'),
            attendeesXML=GetAttendeesXML(attendees)
        )
        resp = self.DoRequest(soapBody)
        self.print('633 resp.text=', resp.text)


class _Attachment:
    def __init__(self, AttachmentId, name, parentExchange):
        print('_Attachment(', AttachmentId, parentExchange)
        self.Filename = name
        self.ID = AttachmentId
        self._parentExchange = parentExchange
        self._content = None

    def _Update(self, getContent=True):
        # sets the filename and content of attachment object

        regExReponse = re.compile(r'<m:ResponseCode>(.+)</m:ResponseCode>')
        regExName = re.compile(r'<t:Name>(.+)</t:Name>')
        regExContent = re.compile(r'<t:Content>(.+)</t:Content>')

        xmlBody = """
                <m:GetAttachment>
                
                    <t:AttachmentShape>
                        
                    </t:AttachmentShape>
                    
                    <m:AttachmentIds>
                        <t:AttachmentId Id="{ID}" />
                    </m:AttachmentIds>
                    
                </m:GetAttachment>""".format(
            ID=self.ID,
        )

        resp = self._parentExchange.DoRequest(xmlBody,
                                              truncatePrint=True)  # the response can be up to 50MB and you dont want to print all that

        responseCode = regExReponse.search(resp.text).group(1)
        if responseCode == 'NoError':  # Handle errors sent by the server
            itemName = regExName.search(resp.text).group(1)
            itemName = itemName.replace(' ', '_')  # remove ' ' chars because dont work on linux
            itemContent = regExContent.search(resp.text).group(1)

            self._content = b64decode(itemContent)
            self.Filename = itemName

    def Read(self):
        if self._content is None:
            self._Update()

        return self._content

    @property
    def Size(self):
        # return size of content in Bytes
        # In theory you could request the size of the attachment from EWS API, or even the hash or changekey
        # but according to this microsoft forum, it is not possible (or at least it does not work as intended)
        # https://social.technet.microsoft.com/Forums/office/en-US/143ab86c-903a-49da-9603-03e65cbd8180/ews-how-to-get-attachment-mime-info-not-content
        return len(self.Read())

    @property
    def Name(self):
        if self.Filename is None:
            self._Update(getContent=False)
        return self.Filename

    def __str__(self):
        return '<Attachment: Name={}>'.format(self.Name)


class ServiceAccount(_ServiceAccountBase):
    def __init__(
            self,
            clientID=None,
            tenantID=None,
            oauthID=None,
            email=None,
            password=None,
            authManager=None
    ):
        self.clientID = clientID
        self.tenantID = tenantID
        self.oauthID = oauthID
        self.email = email
        self.password = password
        self.authManager = authManager

        assert (self.clientID and self.tenantID and self.oauthID and self.authManager) or (
                self.email and self.password), str(self)

    @classmethod
    def Dumper(cls, sa):
        return json.dumps({
            'clientID': sa.clientID,
            'tenantID': sa.tenantID,
            'oauthID': sa.oauthID,
            'email': sa.email,
            'password': sa.password,
            'authManager': 'devices.authManager',  # todo, generalize this
        })

    @classmethod
    def Loader(cls, strng):
        d = json.loads(strng)

        authManager = d.pop('authManager', None)
        if authManager == 'devices.authManager':
            import devices
            authManager = devices.authManager

        ret = cls(
            authManager=authManager,
            **d,
        )
        return ret

    def __str__(self):
        return '<EWS ServiceAccount: clientID={}, tenantID={}, oauthID={}, authManager={}, email={}, password={}>'.format(
            self.clientID[:10] + '...',
            self.tenantID[:10] + '...',
            self.oauthID[:10] + '...',
            self.authManager,
            self.email,
            len(self.password) * '*' if self.password else '***',
        )

    def GetStatus(self):
        try:
            if self.oauthID:
                user = self.authManager.GetUserByID(self.oauthID)
                if user:
                    token = user.GetAccessToken()
                    if token:
                        return 'Authorized'
                    else:
                        return 'Error 633: Token could not be retrieved'
                else:
                    return 'Error 635: User could not be found. OauthID={}'.format(self.oauthID)

            elif self.email:
                ews = EWS(
                    username=self.email,
                    password=self.password,
                )
                ews.UpdateCalendar()
                if 'Connected' in ews.ConnectionStatus:
                    return 'Authorized'
                else:
                    return ews.ConnectionStatus
        except Exception as e:
            return 'Error 650: {}'.format(e)

    def GetType(self):
        return 'Microsoft'

    def GetRoomInterface(self, roomEmail, **kwargs):
        # print('EWS SA.GetRoomInterface(', roomEmail, kwargs)
        if self.oauthID:

            user = self.authManager.GetUserByID(self.oauthID)
            if user is None:
                # ProgramLog(
                #     'EWS ServiceAccount roomEmail="{}" kwargs="{}". '
                #     'No User with ID="{}"'.format(
                #         roomEmail,
                #         kwargs,
                #         self.oauthID
                #     ))
                return

            ews = EWS(
                oauthCallback=user.GetAccessToken,
                impersonation=roomEmail,
                **kwargs
            )
            return ews

        elif self.email:
            ews = EWS(
                username=self.email,
                password=self.password,
                impersonation=roomEmail
            )
            return ews


def GetAttendeesXML(attendeesList):
    if attendeesList:
        attendeesXML = '<t:RequiredAttendees>'
        for a in attendeesList:
            attendeesXML += '''
                                    <t:Attendee>
                                      <t:Mailbox>
                                        <t:EmailAddress>{a}</t:EmailAddress>
                                      </t:Mailbox>
                                    </t:Attendee>
                                '''.format(a=a)
        attendeesXML += '</t:RequiredAttendees>'
        return attendeesXML
    else:
        return ''


def GetCalendarItemDataFromXML(xml):
    '''

    :param xml: str
    :return: list of dicts containing the calendar data
    '''
    ret = []
    for matchCalItem in RE_CAL_ITEM.finditer(xml):
        data = {}

        matchItemId = RE_ITEM_ID.search(matchCalItem.group(0))
        data['ItemId'] = matchItemId.group(1)
        data['ChangeKey'] = matchItemId.group(2)
        data['Subject'] = RE_SUBJECT.search(matchCalItem.group(0)).group(1)
        data['OrganizerName'] = RE_ORGANIZER.search(matchCalItem.group(0)).group(1)

        parentFolderIdMatch = RE_PARENT_FOLDER_ID.search(matchCalItem.group(0))
        if parentFolderIdMatch:
            data['ParentFolderId'] = parentFolderIdMatch.group(1)

        bodyMatch = RE_HTML_BODY.search(matchCalItem.group(0))
        if bodyMatch:
            data['Body'] = bodyMatch.group(1)

        res = RE_HAS_ATTACHMENTS.search(matchCalItem.group(0)).group(1)

        if 'true' in res:
            data['HasAttachments'] = True
        elif 'false' in res:
            data['HasAttachments'] = False
        else:
            data['HasAttachments'] = 'Unknown'

        startTimeString = RE_START_TIME.search(matchCalItem.group(0)).group(1)
        endTimeString = RE_END_TIME.search(matchCalItem.group(0)).group(1)

        startDT = ConvertTimeStringToDatetime(startTimeString)
        endDT = ConvertTimeStringToDatetime(endTimeString)

        calItem = {
            'Start': startDT,
            'End': endDT,
            'Data': data,
        }
        ret.append(calItem)

    return ret


class EWS_BatchManager:
    def __init__(
            self,
            username=None,
            password=None,
            oauthCallback=None,
            debug=True,
            persistentStorageDirectory=None,
    ):
        self.username = username
        self.password = password
        self.oauthCallback = oauthCallback
        self._debug = debug
        self.persistentStorageDirectory = persistentStorageDirectory

        ##############################################
        self.ews_instances = {
            # str(emailAddress): EWS(),
        }

    def print(self, *a, **k):
        if self._debug:
            print(*a, **k)

    def GetEWS(self, email=None):
        email = email.lower() if email else None

        if email and email in self.ews_instances:
            return self.ews_instances[email]
        else:
            newEWS = EWS(
                username=self.username,
                password=self.password,
                oauthCallback=self.oauthCallback,
                impersonation=email,
                debug=self._debug,
                persistentStorage='{}/{}.json'.format(
                    self.persistentStorageDirectory,
                    email,
                ) if self.persistentStorageDirectory and email else None,
            )
            self.ews_instances[email] = newEWS
            return newEWS

    def UpdateCalendarBatch(self, calendars, startDT=None, endDT=None):
        '''

        :param calendars: list of str(emailAddresses)
        :param startDT: datetime object
        :param endDT: datetime object
        :return:
        '''

        # create the batch req
        startDT = startDT or datetime.datetime.now() - datetime.timedelta(days=-1)
        endDT = endDT or startDT + datetime.timedelta(days=2)

        parentFolders = ''
        for cal in calendars:
            parentFolders += '''
                           <t:DistinguishedFolderId Id="calendar">
                               <t:Mailbox>
                                   <t:EmailAddress>{email}</t:EmailAddress>
                               </t:Mailbox>
                           </t:DistinguishedFolderId>
                       '''.format(
                email=cal
            )

        soapBody = ''' 
           <m:FindItem Traversal="Shallow">
               <m:ItemShape>
                   <t:BaseShape>AllProperties</t:BaseShape>
               </m:ItemShape>
               <m:CalendarView 
                   MaxEntriesReturned="100" 
                   StartDate="{startTimestring}" 
                   EndDate="{endTimestring}" 
                   />
               <m:ParentFolderIds>
                    {parentFolders}
               </m:ParentFolderIds>
           </m:FindItem>
        '''.format(
            startTimestring=startDT.isoformat() + 'Z',
            endTimestring=endDT.isoformat() + 'Z',
            parentFolders=parentFolders,
        )
        ews = self.GetEWS()
        resp = ews.DoRequest(soapBody)
        self.print('87 batch resp=', resp.text)

        items = GetCalendarItemDataFromXML(resp.text)

        ret = []
        ews.GetFolderInfo(emails=calendars)
        for item in items:
            email = ews.GetEmailFromFolderID(item['Data']['ParentFolderId'])

            ret.append(
                CalendarItem(
                    startDT=item['Start'],
                    endDT=item['End'],
                    data=item['Data'],
                    parentCalendar=self.GetEWS(email),
                )
            )

        for ews in self.ews_instances.values():
            ews.RegisterCalendarItems(
                calItems=[item for item in ret if item.parentCalendar == ews],
                startDT=startDT,
                endDT=endDT,
            )
            ews.SaveCalendarItemsToFile()

        return ret


if __name__ == '__main__':
    import creds
    import gs_oauth_tools
    import webbrowser

    useOauth = True
    authManager = gs_oauth_tools.AuthManager(
        microsoftClientID=creds.clientID,
        microsoftTenantID=creds.tenantID,
        debug=True
    )

    MY_ID = '3888'

    user = authManager.GetUserByID(MY_ID)
    if user is None:
        d = authManager.CreateNewUser(MY_ID, authType='Microsoft')
        try:
            webbrowser.open(d['verification_uri'])
        except:
            pass
        print('Go to {}'.format(d['verification_uri']))
        print('Enter code "{}"'.format(d['user_code']))
        while authManager.GetUserByID(MY_ID) is None:
            time.sleep(1)
        user = authManager.GetUserByID(MY_ID)
    print('636 user=', user)
    ews = EWS(

        username=None if useOauth else creds.username,
        password=None if useOauth else creds.password,
        impersonation='rnchallwaysignage1@extron.com',
        oauthCallback=None if useOauth is False else user.GetAccessToken,
        debug=True,
        # persistentStorage='test_persistant.json',
    )

    ews.Connected = lambda _, state: print('EWS', state)
    ews.Disconnected = lambda _, state: print('EWS', state)
    ews.NewCalendarItem = lambda cal, item: print('NewCalendarItem(', cal, item)
    ews.CalendarItemChanged = lambda cal, item: print('CalendarItemChanged(', cal, item)
    ews.CalendarItemDeleted = lambda cal, item: print('CalendarItemDeleted(', cal, item)

    ews.UpdateCalendar(
        startDT=datetime.datetime.now().replace(hour=0, minute=0),
        endDT=datetime.datetime.now().replace(hour=23, minute=59)
    )


    # ews.CreateCalendarEvent(
    #     subject='Test Subject ' + time.asctime(),
    #     body='Test Body ' + time.asctime(),
    #     startDT=datetime.datetime.now(),
    #     endDT=datetime.datetime.now() + datetime.timedelta(minutes=15),
    # )

    def TestAttachments():
        while True:
            print('while True')
            ews.UpdateCalendar()
            nowEvents = ews.GetNowCalItems()
            print('nowEvents=', nowEvents)
            for event in nowEvents:
                if 'Test Subject' in event.Get('Subject'):
                    # ews.ChangeEventTime(event, newEndDT=datetime.datetime.now())
                    pass

                if event.HasAttachments():
                    for attach in event.Attachments:
                        attach.Size
                        # open(attach.Name, 'wb').write(attach.Read())
                        print('attach=', attach)

            time.sleep(10)


    # TestAttachments()

    def TestAttendees():
        # ews.CreateCalendarEvent(
        #     subject='Test Subject ' + time.asctime(),
        #     body='Test Body ' + time.asctime(),
        #     startDT=datetime.datetime.now(),
        #     endDT=datetime.datetime.now() + datetime.timedelta(minutes=15),
        #     attendees=['grantm@extrondev.com'],
        # )
        for event in ews.GetNowCalItems():
            if 'gmiller@extron.com' in ews.GetAttendees(event):
                ews.ChangeAttendees(event, removeList=['gmiller@extron.com'])
            else:
                ews.ChangeAttendees(event, addList=['gmiller@extron.com'])


    # TestAttendees()

    def TestBatch():
        import creds
        import time

        batch = EWS_BatchManager(
            username=creds.username,
            password=creds.password,
            persistentStorageDirectory='test',
        )

        calendars = [  # The list of calendars that will be batched together
            'rnchallwaysignage1@extron.com',
            'rnchallwaysignage2@extron.com',
            'rnchallwaysignage3@extron.com',
            'rnclobbysignage@extron.com',
        ]

        calItems = batch.UpdateCalendarBatch(
            calendars=calendars,
        )
        print('len(calitems)=', len(calItems))
        for item in calItems:
            print('item=', item, 'parent=', item.parentCalendar)

        while True:
            time.sleep(5)
            print('while True')

    # TestBatch()
