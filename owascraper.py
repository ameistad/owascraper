#!/usr/bin/python
# -*- coding: utf-8 -*-
"""Outlook Web Access calendar scraper."""
import os
import sys
import re
import datetime

import requests
from bs4 import BeautifulSoup


USERNAME = ''
PASSWORD = ''
URL = 'https://webmail.domain.com'


def main(days):
    with requests.Session() as s:
        login_url = '{}CookieAuth.dll?Logon'.format(URL)
        calendar_url = '{}owa/?ae=Folder&t=IPF.Appointment'.format(URL)

        # Form data
        payload = {
            'username': USERNAME,
            'password': PASSWORD,
            'curl': 'Z2FowaZ2F',
            'flags': '0',
            'forcedownlevel': '0',
            'formdir': '1',
            'trusted': '4'
            }
        
        s.post(login_url, data=payload)
        
        # Get the page ID
        frontpage = BeautifulSoup(s.get(calendar_url).text, 'html.parser')
        page_id = get_page_id(frontpage)

        # Get each day
        what_day = datetime.datetime.now() + datetime.timedelta(days=days)
        calendar_day = '{url}&id={id}&yr={year}&mn={month}&dy={day}'.format(
            url=calendar_url,
            id=page_id,
            year=what_day.year,
            month=what_day.month,
            day=what_day.day)
        
        souped_calendar_day = BeautifulSoup(s.get(calendar_day).text, 'html.parser')
        events_list = [h1.a['title'] for h1 in souped_calendar_day.find_all('h1', 'bld')]

        view_calendar_file = '/Users/andreas/projects/owa_calendar_scraper/view_calendar.md'
        with open(view_calendar_file, 'w') as output_file:
            output_file.write('## OWA Calendar\n')
            output_file.write('#### Showing date: {0}.{1}\n'.format(
                what_day.day, what_day.month))
            for event in events_list:
                output_file.write('{}\n'.format(event))

        bash_command = 'open ' + view_calendar_file
        os.system(bash_command)


def get_page_id(page):
    match_page_id_regex = re.compile('var a_sFldId = "(.*)";')
    for script in page.find_all('script', {'src': False}):
        return match_page_id_regex.search(script.string).group(1)

if __name__ == '__main__':
    try:
        main(int(sys.argv[1]))
    except:
        sys.exit("Invalid argument. Please input number of days from now.")
