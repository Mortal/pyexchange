import argparse
import collections
import datetime
import subprocess

from exchangelib import (
    DELEGATE, Account, Credentials,
    # ServiceAccount,
    # Configuration, NTLM,
    # Build, Version
)
import exchangelib
from exchangelib.services import GetFolder

from exchangelib.util import (
    get_xml_attr, ElementType, xml_to_str,
)
from exchangelib.transport import MNS, TNS
from exchangelib.errors import TransportError
from exchangelib.items import IdOnly
from exchangelib.fields import FieldPath
from exchangelib.folders import DistinguishedFolderId, Calendar
from exchangelib.properties import Mailbox
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone

# Username in WINDOMAIN\username format. Office365 wants usernames in
# PrimarySMTPAddress ('myusername@example.com') format. UPN format is also
# supported.
# import subprocess
# pw = subprocess.check_output(
#     ('pass', 'au'), universal_newlines=True).splitlines()[0]
# credentials = Credentials(username='UNI\\au306325', password=pw)


class ResolveNames(exchangelib.services.ResolveNames):
    def _get_element_container(self, message, name=None):
        assert isinstance(message, ElementType)
        # ResponseClass: See
        # http://msdn.microsoft.com/en-us/library/aa566424(v=EXCHG.140).aspx
        response_class = message.get('ResponseClass')
        # ResponseCode, MessageText: See
        # http://msdn.microsoft.com/en-us/library/aa580757(v=EXCHG.140).aspx
        response_code = get_xml_attr(message, '{%s}ResponseCode' % MNS)
        msg_text = get_xml_attr(message, '{%s}MessageText' % MNS)
        msg_xml = message.find('{%s}MessageXml' % MNS)
        if response_code == 'ErrorNameResolutionMultipleResults':
            container = message.find(name)
            return container
        if response_class == 'Success' and response_code == 'NoError':
            if not name:
                return True
            container = message.find(name)
            if container is None:
                raise TransportError('No %s elements in ResponseMessage (%s)' %
                                     (name, xml_to_str(message)))
            return container
        if response_code == 'NoError':
            return True
        # Raise any non-acceptable errors in the container, or return the
        # container or the acceptable exception instance
        if response_class == 'Warning':
            try:
                return self._raise_errors(code=response_code, text=msg_text,
                                          msg_xml=msg_xml)
            except self.WARNINGS_TO_CATCH_IN_RESPONSE as e:
                return e
        # rspclass == 'Error', or 'Success' and not 'NoError'
        try:
            return self._raise_errors(code=response_code, text=msg_text,
                                      msg_xml=msg_xml)
        except self.ERRORS_TO_CATCH_IN_RESPONSE as e:
            return e


# r = ResolveNames(protocol=account.protocol)
# r.account = account
# results = list(r.call(['5335-327']))
# results = [
#     (mbox.find('{%s}Name' % TNS).text,
#      mbox.find('{%s}EmailAddress' % TNS).text)
#     for r in results
#     for mbox in [r.find('{%s}Mailbox' % TNS)]]
# (name, email), = results
# print(name, email)


# <messages:GetFolder
# xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
# <messages:FolderShape>
# <BaseShape>Default</BaseShape>
# <AdditionalProperties>
# <FieldURI FieldURI="folder:FolderClass"/>
# <FieldURI FieldURI="folder:ParentFolderId"/>
# </AdditionalProperties>
# </messages:FolderShape>
# <messages:FolderIds>
# <DistinguishedFolderId Id="calendar">
# <Mailbox>
# <EmailAddress>
# rmbx.5335327.CS@uni.au.dk</EmailAddress>
# </Mailbox>
# </DistinguishedFolderId>
# </messages:FolderIds>
# </messages:GetFolder>

#    @classmethod
#    def get_distinguished(cls, account, shape=IdOnly):
#        assert shape in SHAPE_CHOICES
#        additional_fields = [FieldPath(field=f) for f in
#                             cls.supported_fields(version=account.version)]
#        folders = []
#        for elem in GetFolder(account=account).call(
#                folders=[cls(account=account)],
#                additional_fields=additional_fields,
#                shape=shape
#        ):
#            if isinstance(elem, Exception):
#                raise elem
#            folder = cls.from_xml(elem=elem, account=account)
#            folder.account = account
#            folders.append(folder)
#        assert len(folders) == 1
#        return folders[0]


class A(DistinguishedFolderId):
    def to_xml(self, version):
        elem = super().to_xml(version)
        elem.append(self.mailbox.to_xml(version))
        return elem

    folder_id = 1


# folders = []
# shape = IdOnly
# additional_fields = [
#     FieldPath(field=f)
#     for f in Calendar.supported_fields(version=account.version)]
#
# a = A(id='calendar')
# a.mailbox = Mailbox(email_address=email)
#
# for elem in GetFolder(account=account).call(
#     folders=[a],
#     additional_fields=additional_fields,
#     shape=IdOnly
# ):
#     folder = Calendar.from_xml(elem=elem, account=account)
#     folder.account = account
#     folders.append(folder)
# print(folders)
# folder, = folders
#
# today = EWSDateTime.from_datetime(
#     datetime.datetime.combine(datetime.date.today(), datetime.time()))
# tomorrow = today + datetime.timedelta(1)
# tz = EWSTimeZone.timezone('Europe/Copenhagen')
# items_for_today = folder.filter(start__range=(
#     tz.localize(today),
#     tz.localize(tomorrow)
# ))  # Filter by a date range
# print(today, tomorrow)
# print(list(items_for_today))

# fid = ('AAMkADk3YjJkNmFlLTI0MWMtNGUyZi04YzNmLTNkZDJmMzVlMTRjZgAuAAAAAA' +
#        'AUGlX4J5VCSa6tjwJYjdcLAQCsIX56+ADZTrONEUqv9COOAAAA8wj/AAA=')
#
#
#
# folder = exchangelib.folders.Calendar(
#     account=account,
#     folder_id=fid)
# folder.refresh()
# print(folder)
# print(list(account.calendar.all()[:10]))


CalendarItem = collections.namedtuple('CalendarItem', 'subject start end')


class ExchangeCalendar:
    def __init__(self, email_address, username, password, calendar_name):
        self.email_address = email_address
        self.username = username
        self.password = password
        self.calendar_name = calendar_name

    @property
    def ews_account(self):
        try:
            return self._cached_ews_account
        except AttributeError:
            pass
        credentials = Credentials(username=self.username,
                                  password=self.password)
        self._cached_ews_account = Account(
            primary_smtp_address=self.email_address,
            credentials=credentials,
            autodiscover=True, access_type=DELEGATE)
        return self._cached_ews_account
        # TODO: No autodiscover:
        # ews_url = account.protocol.service_endpoint
        # ews_auth_type = account.protocol.auth_type
        # primary_smtp_address = account.primary_smtp_address
        # config = Configuration(service_endpoint=ews_url,
        # credentials=credentials, auth_type=ews_auth_type)
        # account = Account(
        #     primary_smtp_address=primary_smtp_address, config=config,
        #     autodiscover=False, access_type=DELEGATE
        # )

    @property
    def calendar_email_address(self):
        try:
            return self._cached_calendar_email_address
        except AttributeError:
            pass
        r = ResolveNames(protocol=self.ews_account.protocol)
        r.account = self.ews_account
        results = list(r.call([self.calendar_name]))
        if len(results) != 1:
            raise ValueError("Got %r results for %r" %
                             (len(results), self.calendar_name))
        results = [
            (mbox.find('{%s}Name' % TNS).text,
             mbox.find('{%s}EmailAddress' % TNS).text)
            for r in results
            for mbox in [r.find('{%s}Mailbox' % TNS)]]
        (name, email), = results
        self._cached_calendar_email_address = email
        return email

    @property
    def ews_calendar(self):
        try:
            return self._cached_ews_calendar
        except AttributeError:
            pass

        folders = []
        account = self.ews_account
        additional_fields = [
            FieldPath(field=f)
            for f in Calendar.supported_fields(version=account.version)]

        a = A(id='calendar')
        a.mailbox = Mailbox(email_address=self.calendar_email_address)

        for elem in GetFolder(account=account).call(
            folders=[a],
            additional_fields=additional_fields,
            shape=IdOnly
        ):
            folder = Calendar.from_xml(elem=elem, account=account)
            folder.account = account
            folders.append(folder)
        if len(folders) != 1:
            raise ValueError("Got %r results, not 1" % len(folders))
        folder, = folders
        self._cached_ews_calendar = folder
        return self._cached_ews_calendar

    def items_for_date(self, date):
        if not isinstance(date, datetime.date):
            raise TypeError(type(date).__name__)
        dt = EWSDateTime.from_datetime(
            datetime.datetime.combine(date, datetime.time()))
        dt2 = dt + datetime.timedelta(1)
        tz = EWSTimeZone.timezone('Europe/Copenhagen')
        items = self.ews_calendar.filter(start__range=(
            tz.localize(dt),
            tz.localize(dt2)
        ))  # Filter by a date range
        return (self.parse_calendar_item(item) for item in items)

    def parse_calendar_item(self, item):
        # CalendarItem(
        #     item_id='AAMkADgyYWUzYzI1LTI5NzktNGQ1YS1hYjgxLTUxYjk1YTZk' +
        #     'MzIyZABGAAAAAAACIa+VsAzWT53TQ+4lbrOvBwBf3o5OX/vTQpW4Gkb+' +
        #     'MZmCAAAABnPcAACJUZ9YoZ9ZRJuUTQUy6X60AAGhpc+jAAA=',
        #     changekey='DwAAABYAAACJUZ9YoZ9ZRJuUTQUy6X60AAGhquEY',
        #     mime_content=b'BEGIN:VCALENDAR\r\n'
        #     b'METHOD:PUBLISH\r\n'
        #     b'PRODID:Microsoft Exchange Server 2010\r\n'
        #     b'VERSION:2.0\r\n'
        #     b'BEGIN:VTIMEZONE\r\n'
        #     b'TZID:Romance Standard Time\r\n'
        #     b'BEGIN:STANDARD\r\n'
        #     b'DTSTART:16010101T030000\r\n'
        #     b'TZOFFSETFROM:+0200\r\n'
        #     b'TZOFFSETTO:+0100\r\n'
        #     b'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=10\r\n'
        #     b'END:STANDARD\r\n'
        #     b'BEGIN:DAYLIGHT\r\n'
        #     b'DTSTART:16010101T020000\r\n'
        #     b'TZOFFSETFROM:+0100\r\n'
        #     b'TZOFFSETTO:+0200\r\n'
        #     b'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3\r\n'
        #     b'END:DAYLIGHT\r\n'
        #     b'END:VTIMEZONE\r\n'
        #     b'BEGIN:VEVENT\r\n'
        #     b'ORGANIZER;CN=Allan Gr\xc3\xb8nlund:MAILTO:jallan@cs.au.dk\r\n'
        #     b'SUMMARY;LANGUAGE=da-DK:Allan Gr\xc3\xb8nlund cbs moede\r\n'
        #     b'DTSTART;TZID=Romance Standard Time:20170814T103000\r\n'
        #     b'DTEND;TZID=Romance Standard Time:20170814T150000\r\n'
        #     b'UID:040000008200E00074C5B7101A82E008000000001D25'
        #     b'9355CF14D301000000000000000\r\n'
        #     b' 010000000174BF105B26F3C4889B29D42AFDC887A\r\n'
        #     b'CLASS:PUBLIC\r\n'
        #     b'PRIORITY:5\r\n'
        #     b'DTSTAMP:20170814T073125Z\r\n'
        #     b'TRANSP:OPAQUE\r\n'
        #     b'STATUS:CONFIRMED\r\n'
        #     b'SEQUENCE:0\r\n'
        #     b'LOCATION;LANGUAGE=da-DK:5335-327 Nygaard '
        #     b'M\xc3\xb8derum (14)\r\n'
        #     b'X-MICROSOFT-CDO-APPT-SEQUENCE:0\r\n'
        #     b'X-MICROSOFT-CDO-OWNERAPPTID:2115531293\r\n'
        #     b'X-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\n'
        #     b'X-MICROSOFT-CDO-INTENDEDSTATUS:BUSY\r\n'
        #     b'X-MICROSOFT-CDO-ALLDAYEVENT:FALSE\r\n'
        #     b'X-MICROSOFT-CDO-IMPORTANCE:1\r\n'
        #     b'X-MICROSOFT-CDO-INSTTYPE:0\r\n'
        #     b'X-MICROSOFT-DISALLOW-COUNTER:FALSE\r\n'
        #     b'END:VEVENT\r\n'
        #     b'END:VCALENDAR\r\n'
        #     b'',
        #     subject='Allan Grønlund cbs moede',
        #     sensitivity='Normal',
        #     text_body=None,
        #     body=None,
        #     attachments=[],
        #     datetime_received=EWSDateTime(2017, 8, 14, 7, 31, 27,
        #                                   tzinfo=<UTC>),
        #     categories=None,
        #     importance='Normal',
        #     is_draft=False,
        #     headers=None,
        #     datetime_sent=EWSDateTime(2017, 8, 14, 7, 31, 27, tzinfo=<UTC>),
        #     datetime_created=EWSDateTime(2017, 8, 14, 7, 31, 27,
        #                                  tzinfo=<UTC>),
        #     reminder_is_set=False,
        #     reminder_due_by=EWSDateTime(2017, 8, 14, 8, 30, tzinfo=<UTC>),
        #     reminder_minutes_before_start=15,
        #     extern_id=None,
        #     last_modified_name='5335-327 Nygaard Møderum (14)',
        #     last_modified_time=EWSDateTime(2017, 8, 14, 7, 31, 27,
        #                                    tzinfo=<UTC>),
        #     conversation_id=ConversationId(
        #         'AAQkADgyYWUzYzI1LTI5NzktNGQ1YS1hYjgxLTUxYjk1YTZkMzIy'
        #         'ZAAQANiiWQyU831OgNnUOC6IRWg=', None),
        #     start=EWSDateTime(2017, 8, 14, 8, 30, tzinfo=<UTC>),
        #     end=EWSDateTime(2017, 8, 14, 13, 0, tzinfo=<UTC>),
        #     is_all_day=False,
        #     legacy_free_busy_status='Busy',
        #     location='5335-327 Nygaard Møderum (14)',
        #     type='Single',
        #     organizer=Mailbox(
        #         'Allan Grønlund', 'jallan@cs.au.dk', 'Mailbox', None),
        #     required_attendees=[
        #         Attendee(
        #             Mailbox(
        #                 'Allan Grønlund', 'jallan@cs.au.dk',
        #                 'Mailbox', None), 'Unknown', None)],
        #     optional_attendees=None,
        #     resources=None,
        #     recurrence=None,
        #     first_occurrence=None,
        #     last_occurrence=None,
        #     modified_occurrences=None,
        #     deleted_occurrences=None)]
        return CalendarItem(item.subject, item.start, item.end)


def parse_date(s):
    return datetime.datetime.strptime(s, '%Y-%m-%d').date()


parser = argparse.ArgumentParser()
parser.add_argument('-e', '--email-address', required=True)
parser.add_argument('-u', '--username', required=True)
parser.add_argument('-p', '--password', required=True)
parser.add_argument('-c', '--calendar-name', required=True)
parser.add_argument('-d', '--date', type=parse_date,
                    default=datetime.date.today())


def main():
    args = parser.parse_args()
    if args.password.startswith('pass:'):
        args.password = subprocess.check_output(
            ('pass', args.password[5:]),
            universal_newlines=True).splitlines()[0]
    args = vars(args)
    date = args.pop('date')
    cal = ExchangeCalendar(**args)
    for i in cal.items_for_date(date):
        print(i)


if __name__ == '__main__':
    main()
