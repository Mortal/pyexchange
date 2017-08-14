import os
import json
import time
import datetime
import requests
import subprocess
from pyexchange import parser, ExchangeCalendar


def to_dict(calendar_item):
    return {'subject': calendar_item.subject,
            'start': calendar_item.start.isoformat(),
            'end': calendar_item.end.isoformat()}


def update(args):
    date = datetime.date.today()
    payload = {'date': date.strftime('%Y-%m-%d'),
               'calendars': {}}
    for c in '5335-395 5335-327'.split():
        args['calendar_name'] = c
        cal = ExchangeCalendar(**args)
        data = list(cal.items_for_date(date))
        payload['calendars'][c] = [to_dict(o) for o in data]
    payload_json = json.dumps(payload)
    url = os.environ.get('LUNCHCLUB_URL',
                         'https://apps.cs.au.dk/lunchclub')
    response = requests.post(
        '%s/calendar/update/' % url,
        dict(token=os.environ['LUNCHCLUB_TOKEN'],
             payload=payload_json))
    if response.status_code >= 300:
        print(response)
        raise Exception("HTTP %s" % response.status_code)


def main():
    args = vars(parser.parse_args())
    args.pop('date')
    if args['password'].startswith('env:'):
        args['password'] = os.environ[args['password'][4:]]
    elif args['password'].startswith('pass:'):
        args['password'] = subprocess.check_output(
            ('pass', args['password'][5:]),
            universal_newlines=True).splitlines()[0]

    while True:
        update(args)
        time.sleep(600)


if __name__ == '__main__':
    main()
