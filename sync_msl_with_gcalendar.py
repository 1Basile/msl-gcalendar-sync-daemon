import argparse
import argparse, textwrap
import copy
import os
import sys
import textwrap
import time
from pprint import pprint
from time import sleep
import datetime
import itertools

import googleapiclient
import httplib2
import rfc3339
from Google import Create_Service
import requests

MSL_ACCESS_TOCKEN = open("msl_token.txt").read()
G_CALENDAR_ID = "q995bvr7o61to88p7qj10830fo@group.calendar.google.com"


class ConstructorError(Exception):
    pass


class MslError(Exception):
    pass


class MslAuthorizationError(MslError):
    pass


class MslConnectivityError(MslError):
    pass


class MslGCalendarCrossEvent():
    """Class for converting Google cloud events to Msl ones,
    and vice-versa."""

    def __init__(self, is_msl_exam=False, is_msl_class=False, is_gcalendar=False, event=None):
        self.__GColorNum_to_ColorName = {'3': 'Grape', '1': 'Lavender', '8': 'Graphite',
                                         '9': 'Blueberry', '7': 'Peacock', '10': 'Basil',
                                         '2': 'Banana', '6': 'Tangerine', '4': 'Flamingo',
                                         '11': 'Tomato'}
        self.__ColorName_to_GColorNum = {v: k for k, v in self.__GColorNum_to_ColorName.items()}
        self.__MslColorNum_to_ColorName = {'4': 'Flamingo', '8': 'Lavender', '14': 'Grape',
                                           '1': 'Tangerine', '2': 'Banana', '9': 'Peacock',
                                           '3': 'Blueberry', '6': 'Graphite', '13': 'Basil',
                                           '12': 'Tomato'}
        self.__ColorName_to_MslColorNum = {v: k for k, v in self.__MslColorNum_to_ColorName.items()}

        self.color = None
        self.is_recurrence = "recurrence" in event
        self.timeZone = None
        self.created = None  # With no .000Z at the end
        self.end = None  # With no :00Z at the end
        self.recurrence = []
        self.start = None  # With no :00Z at the end
        self.date_class_ends = None
        self.title = ''
        self.description = ""
        self.event_id = None

        if is_gcalendar and not is_msl_class:
            self.event_id = event['id']
            try:
                self.timeZone = event['start']['timeZone']
            except:
                self.timeZone = 'Europe/Kiev'

            try:
                self.color = self.__GColorNum_to_ColorName[event['colorId']]
            except:
                pass

            self.created = event['created'].rsplit('.', 1)[0]

            start = event['start']['dateTime'].rsplit(':', 1)[0] + ':00Z'
            end = event['end']['dateTime'].rsplit(':', 1)[0] + ':00Z'

            if self.is_recurrence:
                for recurrence in event["recurrence"]:
                    date_ = event['end']['dateTime'].split('T', 1)[0].split('-')
                    self.recurrence.append(dict(days=list(day for day in recurrence.rsplit('BYDAY=', 1)[1].split(',')),
                                                is_2_rotation_week="INTERVAL=2" in recurrence,
                                                date_class_ends=recurrence.rsplit('UNTIL=', 1)[1].split(';', 1)[
                                                    0] if 'UNTIL=' in recurrence else None,
                                                # ['RRULE:FREQ=WEEKLY;WKST=MO;UNTIL=20211231T215959Z;BYDAY=MO']
                                                week_parity=datetime.date(
                                                    *list([int(i) for i in date_])).isocalendar()[
                                                                1] % 2 + 1 if "INTERVAL=2" in recurrence else 0,
                                                end=end,
                                                start=start))

            else:
                self.end = end
                self.start = start

            if 'summary' in event:
                self.title = event['summary']
            if 'description' in event:
                self.description = event['description']

        elif is_msl_class and not is_gcalendar:
            num_to_day = {'2': 'MO', '4': 'TU', '8': 'WE', '16': 'TH', '32': 'FR', '64': 'SA', '1': 'SU'}
            name_to_abriviature = {'MO': 'Mon', 'TU': 'Tue', 'WE': 'Wed',
                                   'TH': 'Thu', 'FR': 'Fri', 'SA': 'Sat', 'SU': 'Sun'}
            find_first_week_day_after = lambda date, day: date + datetime.timedelta(days=(day - date.weekday() + 7) % 7)
            self.is_recurrence = True

            try:  # Msl gives more colors than Gooogle Calendar
                self.color = self.__MslColorNum_to_ColorName[event['subject_color']]
            except:
                pass

            self.title = event['subject_name'] + '/' + event['module']
            for one_time in event['times']:
                days = [num_to_day[str(one_time['days'])]]
                week_parity = one_time['rotation_week']
                is_2_rotation_week = week_parity != 0

                __ = event['subj_start_date'].split('-') + one_time['start_time'].split(":")
                _ = event['subj_start_date'].split('-') + one_time['end_time'].split(":")  # subj_end_date

                subj_end_date_list = event['subj_end_date'].split('-') + ['T'] + ['21', '59', '59', 'Z']

                _start_time = datetime.datetime(*[int(i) for i in __], 000000)
                start_time = find_first_week_day_after(_start_time, time.strptime(name_to_abriviature[days[0]], '%a')
                                                       .tm_wday)

                _end_time = datetime.datetime(*[int(i) for i in _], 000000)
                end_time = find_first_week_day_after(_end_time, time.strptime(name_to_abriviature[days[0]], '%a')
                                                     .tm_wday)

                if week_parity == '1':  # if week has - parity
                    start_time = start_time + datetime.timedelta(weeks=1)
                if not self.timeZone:
                    start_time = start_time - datetime.timedelta(hours=3)
                    end_time = end_time - datetime.timedelta(hours=3)
                start = rfc3339.rfc3339(start_time).rsplit('+', 1)[0].rsplit(':', 1)[0] + ':00Z'
                end = rfc3339.rfc3339(end_time).rsplit('+', 1)[0].rsplit(':', 1)[0] + ':00Z'

                self.recurrence.append(dict(days=days,
                                            is_2_rotation_week=is_2_rotation_week,
                                            week_parity=week_parity,
                                            end=end,
                                            date_class_ends=''.join(subj_end_date_list),
                                            start=start))

        elif is_msl_exam and not is_msl_class:
            self.is_recurrence = False
            self.title = "Exam: " + event['subject_name'] + '/'
            if event['module']:
                self.title += event['module']

            try:  # Msl gives more colors than Gooogle Calendar
                self.color = self.__MslColorNum_to_ColorName[event['subject_color']]
            except:
                pass

            _ = event['date'].split("T", 1)[0].split('-') + event['date'].split("T", 1)[1].split(":")
            _start_time = datetime.datetime(*[int(i) for i in _], 000000)

            start_time = _start_time
            end_time = _start_time + datetime.timedelta(minutes=int(event['duration']))

            if not self.timeZone:
                start_time = start_time - datetime.timedelta(hours=3)
                end_time = end_time - datetime.timedelta(hours=3)

            self.start = rfc3339.rfc3339(start_time).rsplit('+', 1)[0].rsplit(':', 1)[0] + ':00Z'
            self.end = rfc3339.rfc3339(end_time).rsplit('+', 1)[0].rsplit(':', 1)[0] + ':00Z'

        else:
            raise ConstructorError("Create object with one of classmethods, not explicitly.")

    def to_gcalendar_events(self):
        events = []
        if self.is_recurrence:
            for recurrence in self.recurrence:
                event = {
                    'summary': self.title,
                    'start': {
                        'dateTime': recurrence['start'],
                        'timeZone': 'Europe/Kiev' if not self.timeZone else self.timeZone
                    },
                    'end': {
                        'dateTime': recurrence['end'],
                        'timeZone': 'Europe/Kiev' if not self.timeZone else self.timeZone
                    },
                    'recurrence':
                        [
                            'RRULE:FREQ=WEEKLY;WKST=MO;{0}{1}BYDAY={2}'.format(
                                'INTERVAL=2;' if recurrence["is_2_rotation_week"] else '',
                                'UNTIL=' + recurrence['date_class_ends'] + ';' if recurrence['date_class_ends'] else '',
                                ','.join(recurrence["days"]))
                        ]
                }
                if self.color:
                    event['colorId'] = self.__ColorName_to_GColorNum[self.color]
                if self.description:
                    event['description'] = self.description
                if self.event_id:
                    event['id'] = self.event_id
                events.append(event)
            return events
        else:
            event = {
                'summary': self.title,
                'description': self.description,
                'start': {
                    'dateTime': self.start,
                    'timeZone': 'Europe/Kiev' if not self.timeZone else self.timeZone
                },
                'end': {
                    'dateTime': self.end,
                    'timeZone': 'Europe/Kiev' if not self.timeZone else self.timeZone
                }
            }
            if self.color:
                event['colorId'] = self.__ColorName_to_GColorNum[self.color]
            if self.description:
                event['description'] = self.description
            if self.event_id:
                event['id'] = self.event_id

            return [event, ]

    @classmethod
    def from_msl_class(cls, event: str):
        return cls(event=event, is_msl_class=True)

    @classmethod
    def from_msl_exam(cls, event: str):
        return cls(event=event, is_msl_exam=True)

    @classmethod
    def from_gcalendar_event(cls, event: str):
        return cls(event=event, is_gcalendar=True)


class MyGoogleCalendar():
    """Class for handling Google Calendar events adding/deleting/editing them."""

    def __init__(self, calendarId=None):
        _CLIENT_SECRET_FILE = 'cred.json'
        _API_NAME = 'calendar'
        _API_VERSION = 'v3'
        _SCOPES = ['https://www.googleapis.com/auth/calendar']
        _SERVICE_ACCOUNT_FILE = 'cred.json'

        self.calendarId = calendarId

        self.service = Create_Service(_CLIENT_SECRET_FILE, _API_NAME, _API_VERSION, _SCOPES)

    def get_events(self, calendarId=None):
        """Method return all events in exact :calendar:."""
        if calendarId is None:
            calendarId = self.calendarId

        result = []

        page_token = None
        while True:
            events = self.service.events().list(calendarId=self.calendarId, pageToken=page_token).execute()
            for event in events['items']:
                result.append(event)
            page_token = events.get('nextPageToken')
            if not page_token:
                break

        return result

    def create_events(self, events, calendarId=None):
        if calendarId is None:
            calendarId = self.calendarId
        res = []
        for event in events:
            if "id" in event:  # remove id
                event.pop("id")
            res.append(self.service.events().insert(calendarId=calendarId, body=event).execute())
        return res

    def delete_events(self, events, calendarId=None):
        if calendarId is None:
            calendarId = self.calendarId
        res = []
        for event in events:
            if 'id' in event:
                res.append(self.service.events().delete(calendarId=calendarId, eventId=event["id"]).execute())
        return res


class MyMslCalendar():
    """Class for accessing My Study Life schedule."""

    def __init__(self, accessToken=None):
        self.__API_URL = "https://api.mystudylife.com/v6.1"
        if accessToken is None:
            self.__accessToken = MSL_ACCESS_TOCKEN
        else:
            self.__accessToken = accessToken

        self.__headers = {"Accept": "application/json", "Authorization": f"Bearer {self.__accessToken}"}

        self.test_connectivity()

    def _get_data(self):
        response = requests.get(f"{self.__API_URL}/data", headers=self.__headers).json()

        if 'error' in response and \
                response['error_message'] == 'API calls quota exceeded! maximum admitted 2 per Second.':
            sleep(2)
            response = requests.get(f"{self.__API_URL}/data", headers=self.__headers).json()

        return response

    def test_connectivity(self):
        _ = self._get_data()
        if 'error' in _:
            if _['error'] == 'not_authorized':
                raise MslAuthorizationError(_['error_message'])
            else:
                raise MslConnectivityError(_['error_message'])

    def get_classes_ext(self):
        """
        Classes respond fields are modified:
            1) added filed color from appropriate subject
            2) added filed 'subject_name' from appropriate subject
        """
        classes = self._get_data()["classes"]
        academic_years = {
            i["guid"]: {
                'start_date': i["start_date"],
                'end_date': i["end_date"],
            } for i in self._get_data()["academic_years"]
        }
        terms = {}
        for academic_year in self._get_data()["academic_years"]:
            _terms = {
                i["guid"]: {
                    'start_date': i["start_date"],
                    'end_date': i["end_date"],
                    'name': i['name']
                } for i in academic_year['terms']
            }
            terms = {**terms, **_terms}

        subjects_dict = {
            i["guid"]: {
                'color': i["color"],
                'name': i["name"],
                'year_guid': i['year_guid'],
                'term_guid': i['term_guid']
            } for i in self.get_subjects()}

        for _class in classes:
            _subj = subjects_dict[_class['subject_guid']]
            _class['subject_color'] = _subj['color']
            _class['subject_name'] = _subj['name']
            if _subj['term_guid']:
                _term = terms[_subj['term_guid']]
                _class['subj_start_date'] = _term['start_date']
                _class['subj_end_date'] = _term['end_date']
            else:
                _year = academic_years[_subj['term_guid']]
                _class['subj_start_date'] = _year['start_date']
                _class['subj_end_date'] = _year['end_date']

        return classes

    def get_subjects(self):
        return self._get_data()["subjects"]

    def get_exams_ext(self):
        """
        Classes respond fields are modified:
            1) added filed color from appropriate subject
            2) added filed 'subject_name' from appropriate subject
        """
        exams = self._get_data()["exams"]
        exams_dict = {
            i["guid"]: {
                'color': i["color"],
                'name': i["name"],
                'year_guid': i['year_guid'],
                'term_guid': i['term_guid']
            } for i in self.get_subjects()}

        for _exam in exams:
            _subj = exams_dict[_exam['subject_guid']]
            _exam['subject_color'] = _subj['color']
            _exam['subject_name'] = _subj['name']

        return exams


# Disable
def blockPrint():
    sys.stdout = open(os.devnull, 'w')


# Restore
def enablePrint():
    sys.stdout = sys.__stdout__


def get_wrong_g_events(g_events, msl_events):
    """Return list of google calendar events, of thous, whose Msl equivalent does not exist."""
    diff = []
    listed = []
    for i in g_events:  # Remove 'id' key
        cp_i = copy.deepcopy(i)
        cp_i.pop('id')
        if (cp_i not in msl_events) or (cp_i in listed):
            diff.append(i)
        listed.append(cp_i)
    return diff


def get_missing_msl_events(g_events, msl_events):
    """Return list of missing msl classes, thous, whose google calendar events does not exist."""
    diff = []
    un_id_g_events = []
    for i in g_events:
        cp_i = copy.deepcopy(i)
        cp_i.pop('id')
        un_id_g_events.append(cp_i)

    for i in msl_events:
        if i not in un_id_g_events:
            diff.append(i)

    return diff


def sync_data():
    google_calendar = MyGoogleCalendar(calendarId=G_CALENDAR_ID)
    msl_calendar = MyMslCalendar()

    __uni_g_events = [MslGCalendarCrossEvent.from_gcalendar_event(event) for event in google_calendar.get_events()]
    __g_from_uni_g_events_flatten = list(itertools.chain(*[event.to_gcalendar_events() for event in __uni_g_events]))
    g_events = __g_from_uni_g_events_flatten

    __uni_msl_events = [MslGCalendarCrossEvent.from_msl_class(event) for event in msl_calendar.get_classes_ext()]
    __g_from_uni_msl_events_flatten = list(
        itertools.chain(*[event.to_gcalendar_events() for event in __uni_msl_events]))
    msl_events = __g_from_uni_msl_events_flatten

    __uni_msl_exams = [MslGCalendarCrossEvent.from_msl_exam(event) for event in msl_calendar.get_exams_ext()]
    __g_from_uni_msl_exams_flatten = list(itertools.chain(*[event.to_gcalendar_events() for event in __uni_msl_exams]))
    msl_exams = __g_from_uni_msl_exams_flatten

    # delete old/wrong events
    wrong_g_events = get_wrong_g_events(g_events, msl_events + msl_exams)
    google_calendar.delete_events(wrong_g_events)

    print("Deleted from wrong GCalendar events:")
    pprint(wrong_g_events)

    # get missing msl events
    missing_msl_events = get_missing_msl_events(g_events, msl_events + msl_exams)
    google_calendar.create_events(missing_msl_events)
    print("New created/updated Msl events to add to GCalendar:")
    pprint(missing_msl_events)


def main():

    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
        description='Script sync MyStudyLife account classes with GCalendar.',
        epilog=textwrap.dedent('''\
            Exit statues:
            -------------------------
                1) No internet connection.
                2) Set wrong Msl token.
                3) Some other Msl error.
                4) Some Google side error.
                '''))

    parser.add_argument('--show-creds', help='Show MSL token and Google Calendar ID.', action='store_true')
    parser.add_argument('-v', help='Show verbose output.', action='store_true')

    args = parser.parse_args()

    if args.show_creds:
        print("My Study Life token:", MSL_ACCESS_TOCKEN)
        print("Google Calendar ID:", G_CALENDAR_ID)
        sys.exit(0)

    if args.v:
        enablePrint()
    else:
        blockPrint()

    try:
        sync_data()
        sys.exit(0)

    except httplib2.error.ServerNotFoundError as e:  # No internet connection.
        sys.stdout = sys.__stderr__
        print("ERROR:", e.__str__())
        sys.exit(1)
    except MslAuthorizationError as e:  # Wrong Msl token
        sys.stdout = sys.__stderr__
        print("ERROR: Msl service.", e.__str__())
        sys.exit(3)
    except MslError as e:  # Some other Msl error
        sys.stdout = sys.__stderr__
        print("ERROR: Msl service.", e.__str__())
        sys.exit(4)
    except googleapiclient.errors.HttpError as e:  # Some Google side error.
        sys.stdout = sys.__stderr__
        print("ERROR: Google calendar service.", e.__str__())
        sys.exit(5)
    except Exception:
        sys.stdout = sys.__stderr__
        print("ERROR: Undefined error.")
        sys.exit(6)


if __name__ == '__main__':
    main()
