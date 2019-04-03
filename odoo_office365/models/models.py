# -*- coding: utf-8 -*-

import logging
import re
from urllib.parse import urlparse
from selenium import webdriver
import requests
import json
from datetime import datetime
import time
from datetime import timedelta
import os

from odoo import fields, models, api, osv
from openerp.exceptions import ValidationError
from openerp.osv import osv
from odoo import exceptions, _
from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools

_logger = logging.getLogger(__name__)
_image_dataurl = re.compile(r'(data:image/[a-z]+?);base64,([a-z0-9+/]{3,}=*)([\'"])', re.I)
root_path = os.path.dirname(os.path.abspath(__file__))


class OfficeSettings(models.Model):
    """
    This class separates one time office 365 settings from Token generation settings
    """

    _name = "office.settings"

    field_name = fields.Char('Office365')
    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret = fields.Char('Secret')
    microsoft_email = fields.Char(String='Microsoft Account Email')
    microsoft_password = fields.Char(String='Microsoft Account Password')
    login_url = None
    stay_signed_in = False

    @api.one
    def sync_data(self):

        """
            This function checks the connection the user account
        """

        try:

            self.login_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=%s&redirect_uri=%s&response_type=code&scope=openid+offline_access+Group.ReadWrite.All+Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+User.ReadWrite+Tasks.ReadWrite' \
                             % (self.client_id, self.redirect_url)
            path_to_chromedriver = os.path.join(root_path, 'chromedriver')

            options = webdriver.ChromeOptions()
            options.add_argument('--disable-extensions')
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')

            driver = webdriver.Chrome(executable_path=path_to_chromedriver, chrome_options=options)
            driver.get(self.login_url)

            driver.find_element_by_id('i0116').send_keys(self.microsoft_email)
            driver.find_element_by_id('idSIButton9').click()
            time.sleep(5)
            driver.find_element_by_id('i0118').send_keys(self.microsoft_password)
            driver.find_element_by_id('idSIButton9').click()
            time.sleep(1)

            try:
                driver.find_element_by_id('idBtn_Accept').click()
            except Exception as e:
                _logger.exception(e)

            try:
                driver.find_element_by_id('idSIButton9').click()
            except Exception as e:
                _logger.exception(e)

            code = urlparse(driver.current_url)[4].split('=')[1]
            driver.close()

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data='grant_type=authorization_code&code=' + code + '&redirect_uri=' + self.redirect_url + '&client_id=' + self.client_id + '&client_secret=' + self.secret
                , headers=header).content
            response = json.loads((str(response)[2:])[:-1])
            self.env.user.redirect_url = self.redirect_url
            self.env.user.client_id = self.client_id
            self.env.user.secret = self.secret
            self.env.user.token = response['access_token']
            self.env.user.refresh_token = response['refresh_token']
            self.env.user.expires_in = int(round(time.time() * 1000))
            self.env.user.code = code
            self.env.user.redirect_url = self.redirect_url
            self.env.user.client_id = self.client_id
            self.env.user.secret = self.secret

            response = json.loads((requests.get(
                'https://graph.microsoft.com/v1.0/me',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content.decode('utf-8')))
            self.env.user.office365_email = response['userPrincipalName']
            self.env.user.office365_id_address = 'outlook_' + response['id'].upper() + '@outlook.com'
            self.env.cr.commit()

            office_credentials = self.env['office.credentials'].search([])

            if not office_credentials:
                self.env['office.credentials'].create({'redirect_url': self.redirect_url,
                                                        'code': code,
                                                        'client_id': self.client_id,
                                                        'secret': self.secret,
                                                        'microsoft_email': self.microsoft_email,
                                                        'microsoft_password': self.microsoft_password
                                                        })
            self.env.cr.commit()

        except Exception as e:
            raise ValidationError(_(str(e)))
        raise osv.except_osv(_("Success!"), (_("Token Generated!")))


class Office365UserSettings(models.Model):
    """
    This class facilitates the users other than admin to enter office 365 credential

    params _name :Is a Custom Model Name
    """

    _name = 'office.usersettings'

    login_url = fields.Char('Login URL', compute='_compute_url', readonly=True)
    code = fields.Char('code')
    field_name = fields.Char('office')
    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret = fields.Char('Secret')
    microsoft_email = fields.Char(String='Microsoft Account Email')
    microsoft_password = fields.Char(String='Microsoft Account Password')

    @api.one
    def _compute_url(self):
        """
         if the user provide the office 365 credentials then this function generates a URL using this credential to generate oken

        """
        settings = self.env['office.credentials'].search([])
        settings = settings[0] if settings else settings
        if settings:
            self.login_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=%s&redirect_uri=%s&response_type=code&scope=openid+offline_access+Group.ReadWrite.All+Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+User.ReadWrite+Tasks.ReadWrite' \
                             % (settings.client_id, settings.redirect_url)

    @api.one
    def test_connectiom(self):
        """
        this function tests credential and generates token

        """
        try:
            settings = self.env['office.credentials'].search([])
            settings = settings[0] if settings else settings

            if not settings.client_id or not settings.redirect_url or not settings.secret:
                raise osv.except_osv(_("Error!"), (_("Please ask admin to add Office365 settings!")))

            path_to_chromedriver = os.path.join(root_path, 'chromedriver')

            options = webdriver.ChromeOptions()
            options.add_argument('--disable-extensions')
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')

            driver = webdriver.Chrome(executable_path=path_to_chromedriver, chrome_options=options)
            driver.get(self.login_url)

            driver.find_element_by_id('i0116').send_keys(settings.microsoft_email)
            driver.find_element_by_id('idSIButton9').click()
            time.sleep(5)
            driver.find_element_by_id('i0118').send_keys(settings.microsoft_password)
            driver.find_element_by_id('idSIButton9').click()
            time.sleep(1)

            try:
                driver.find_element_by_id('idBtn_Accept').click()
            except Exception as e:
                _logger.exception(e)

            try:
                driver.find_element_by_id('idSIButton9').click()
            except Exception as e:
                _logger.exception(e)

            code = urlparse(driver.current_url)[4].split('=')[1]
            driver.close()

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data='grant_type=authorization_code&code=' + code + '&redirect_uri=' + settings.redirect_url + '&client_id=' + settings.client_id + '&client_secret=' + settings.secret
                , headers=header).content
            response = json.loads((str(response)[2:])[:-1])
            self.env.user.redirect_url = settings.redirect_url
            self.env.user.client_id = settings.client_id
            self.env.user.secret = settings.secret
            self.env.user.token = response['access_token']
            self.env.user.refresh_token = response['refresh_token']
            self.env.user.expires_in = int(round(time.time() * 1000))
            self.env.user.code = code

            response = json.loads((requests.get(
                'https://graph.microsoft.com/v1.0/me',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content.decode('utf-8')))
            self.env.user.office365_email = response['userPrincipalName']
            self.env.user.office365_id_address = 'outlook_' + response['id'].upper() + '@outlook.com'
            self.env.cr.commit()

        except Exception as e:
            raise ValidationError(_(str(e)))

        raise osv.except_osv(_("Success!"), (_("Token Generated!")))


class CustomUser(models.Model):
    """
    This class adds functionality to user for Office365 Integration
    """
    _inherit = 'res.users'

    login_url = fields.Char('Login URL', compute='_compute_url', readonly=True)
    code = fields.Char('code')
    token = fields.Char('Token', readonly=True)
    refresh_token = fields.Char('Refresh Token', readonly=True)
    expires_in = fields.Char('Expires IN', readonly=True)
    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret = fields.Char('Secret')
    office365_email = fields.Char('Office365 Email Address', readonly=True)
    office365_id_address = fields.Char('Office365 Id Address', readonly=True)
    send_mail_flag = fields.Boolean(string='Send messages using office365 Mail', default=True)
    is_task_sync_on = fields.Boolean('is sync in progress', default=False)

    @api.one
    def _compute_url(self):
        """
        this function creates a url. By hitting this URL creates a code that is require to generate token. That token will be sent with every API request

        :return:
        """
        self.login_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=%s&redirect_uri=%s&response_type=code&scope=openid+offline_access+Group.ReadWrite.All+Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+User.ReadWrite+Tasks.ReadWrite' % (
            self.client_id, self.redirect_url)

    @api.one
    def test_connectiom(self):
        """
        This function generates token using code generated using above login URL
        :return:
        """
        try:
            if not self.client_id or not self.redirect_url or not self.secret:
                raise osv.except_osv(_("Error!"), (_(
                    "Please go to settings--> users & companies--> Office365 and enter credentials and press activate !")))

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            if self.refresh_token and self.expires_in:
                response = requests.post(
                    'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                    data='grant_type=refresh_token&refresh_token=' + self.refresh_token + '&redirect_uri=' + self.redirect_url + '&client_id=' + self.client_id + '&client_secret=' + self.secret
                    , headers=header).content
            else:
                response = requests.post(
                    'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                    data='grant_type=authorization_code&code=' + self.code + '&redirect_uri=' + self.redirect_url + '&client_id=' + self.client_id + '&client_secret=' + self.secret
                    , headers=header).content

            response = json.loads((str(response)[2:])[:-1])
            if 'access_token' not in response:
                response["error_description"] = response["error_description"].replace("\\r\\n", " ")
                raise osv.except_osv(_("Error!"), (_(response["error"] + " " + response["error_description"])))
            else:
                self.token = response['access_token']
                self.refresh_token = response['refresh_token']
                self.expires_in = int(round(time.time() * 1000))

                response = json.loads((requests.get(
                    'https://graph.microsoft.com/v1.0/me',
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.token),
                        'Accept': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content.decode('utf-8')))
                self.office365_email = response['userPrincipalName']
                self.office365_id_address = 'outlook_' + response['id'].upper() + '@outlook.com'
                self.env.cr.commit()

        except Exception as e:
            raise ValidationError(_(str(e)))

        raise osv.except_osv(_("Success!"), (_("Token Generated!")))

    @api.model
    def auto_import_calendar(self):
        self.import_calendar()

    @api.model
    def auto_export_calendar(self):
        self.export_calendar()

    # @api.one
    def import_calendar(self):
        """
        this function imports Office 365  Calendar to Odoo Calendar

        :return:
        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/events',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'value' not in json.loads((response.decode('utf-8'))).keys():
                raise osv.except_osv(response)
            events = json.loads((response.decode('utf-8')))['value']
            for event in events:
                odoo_meeting = self.env['calendar.event'].search([("office_id", "=", event['id'])])
                if odoo_meeting:
                    odoo_meeting.unlink()
                    self.env.cr.commit()

                odoo_event = self.env['calendar.event'].create({
                    'office_id': event['id'],
                    'name': event['subject'],
                    'location': (event['location']['address']['city'] + ', ' + event['location']['address'][
                        'countryOrRegion']) if 'address' in event['location'] and 'city' in event['location'][
                        'address'].keys() else "",
                    'start': datetime.strptime(event['start']['dateTime'][:-8], '%Y-%m-%dT%H:%M:%S').strftime(
                        '%Y-%m-%d %H:%m:%S'),
                    'stop': datetime.strptime(event['end']['dateTime'][:-8], '%Y-%m-%dT%H:%M:%S').strftime(
                        '%Y-%m-%d %H:%m:%S'),
                    'allday': event['isAllDay'],
                    'show_as': event['showAs'],
                    'recurrency': True if event['recurrence'] else False,
                    'end_type': 'end_date' if event['recurrence'] else "",
                    'rrule_type': event['recurrence']['pattern']['type'].replace('absolute', '').lower() if event[
                        'recurrence'] else "",
                    'count': event['recurrence']['range']['numberOfOccurrences'] if event['recurrence'] else "",
                    'final_date': datetime.strptime(event['recurrence']['range']['endDate'], '%Y-%m-%d').strftime(
                        '%Y-%m-%d') if event['recurrence'] else None,
                    'mo': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'monday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                    'tu': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'tuesday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                    'we': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'wednesday' in event['recurrence']['pattern'][
                                      'daysOfWeek'] else False,
                    'th': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'thursday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                    'fr': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'friday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                    'sa': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'saturday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                    'su': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                        'pattern'].keys() and 'sunday' in event['recurrence']['pattern']['daysOfWeek'] else False,
                })
                partner_ids = []
                attendee_ids = []
                for attendee in event['attendees']:
                    partner = self.env['res.partner'].search([('email', "=", attendee['emailAddress']['address'])])
                    if not partner:
                        partner = self.env['res.partner'].create({
                            'name': attendee['emailAddress']['name'],
                            'email': attendee['emailAddress']['address'],
                        })
                    partner_ids.append(partner[0].id)
                    odoo_attendee = self.env['calendar.attendee'].create({
                        'partner_id': partner[0].id,
                        'event_id': odoo_event.id,
                        'email': attendee['emailAddress']['address'],
                        'common_name': attendee['emailAddress']['name'],

                    })
                    attendee_ids.append(odoo_attendee.id)
                if not event['attendees']:
                    odoo_attendee = self.env['calendar.attendee'].create({
                        'partner_id': self.env.user.partner_id.id,
                        'event_id': odoo_event.id,
                        'email': self.env.user.partner_id.email,
                        'common_name': self.env.user.partner_id.name,

                    })
                    attendee_ids.append(odoo_attendee.id)
                    partner_ids.append(self.env.user.partner_id.id)
                odoo_event.write({
                    'attendee_ids': [[6, 0, attendee_ids]],
                    'partner_ids': [[6, 0, partner_ids]]
                })
                self.env.cr.commit()



        except Exception as e:
            raise ValidationError(_(str(e)))

        # raise osv.except_osv(_("Success!"), (_(" Sync Successfully !")))

    # @api.one
    def export_calendar(self):
        """
        this function export  odoo calendar event  to office 365 Calendar

        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            header = {
                'Authorization': 'Bearer {0}'.format(self.env.user.token),
                'Content-Type': 'application/json'
            }
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/calendars',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'value' not in json.loads((response.decode('utf-8'))).keys():
                raise osv.except_osv(("Access Token Expired!"), (" Please Regenerate Access Token !"))
            calendars = json.loads((response.decode('utf-8')))['value']
            calendar_id = calendars[0]['id']
            meetings = self.env['calendar.event'].search([("office_id", "=", False)])
            added_meetings = self.env['calendar.event'].search([("office_id", "!=", False)])

            added = []
            for meeting in meetings:
                temp = meeting
                id = str(meeting.id).split('-')[0]
                metngs = [meeting for meeting in meetings if id in str(meeting.id)]
                index = len(metngs)
                meeting = metngs[index - 1]
                if meeting.start is not None:
                    metting_start = meeting.start.strftime('%Y-%m-%d T %H:%M:%S') if meeting.start else meeting.start
                else:
                    metting_start = None

                payload = {
                    "subject": meeting.name,
                    "attendees": self.getAttendee(meeting.attendee_ids),
                    'reminderMinutesBeforeStart': self.getTime(meeting.alarm_ids),
                    "start": {
                        "dateTime": meeting.start.strftime('%Y-%m-%d T %H:%M:%S') if meeting.start else meeting.start,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": meeting.stop.strftime('%Y-%m-%d T %H:%M:%S') if meeting.stop else meeting.stop,
                        "timeZone": "UTC"
                    },
                    "showAs": meeting.show_as,
                    "location": {
                        "displayName": meeting.location if meeting.location else "",
                    },

                }
                if meeting.recurrency:
                    payload.update({"recurrence": {
                        "pattern": {
                            "daysOfWeek": self.getdays(meeting),
                            "type": (
                                        'Absolute' if meeting.rrule_type != "weekly" and meeting.rrule_type != "daily" else "") + meeting.rrule_type,
                            "interval": meeting.interval,
                            "month": int(meeting.start.month),  # meeting.start[5] + meeting.start[6]),
                            "dayOfMonth": int(meeting.start.day),  # meeting.start[8] + meeting.start[9]),
                            "firstDayOfWeek": "sunday",
                            # "index": "first"
                        },
                        "range": {
                            "type": "endDate",
                            "startDate": str(meeting.start.year + "-" + meeting.start.month + "-" + meeting.start.day),
                            "endDate": str(meeting.final_date),
                            "recurrenceTimeZone": "UTC",
                            "numberOfOccurrences": meeting.count,
                        }
                    }})
                if meeting.name not in added:
                    response = requests.post(
                        'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events',
                        headers=header, data=json.dumps(payload)).content

                    temp.write({
                        'office_id': json.dumps((response.decode('utf-8')))[0]
                    })
                    self.env.cr.commit()
                    if meeting.recurrency:
                        added.append(meeting.name)

            added = []
            for meeting in added_meetings:
                id = str(meeting.id).split('-')[0]
                metngs = [meeting for meeting in added_meetings if id in str(meeting.id)]
                index = len(metngs)
                meeting = metngs[index - 1]

                payload = {
                    "subject": meeting.name,
                    "attendees": self.getAttendee(meeting.attendee_ids),
                    'reminderMinutesBeforeStart': self.getTime(meeting.alarm_ids),
                    "start": {
                        "dateTime": meeting.start.strftime('%Y-%m-%d T %H:%M:%S') if meeting.start else meeting.start,
                        "timeZone": "UTC"
                    },

                    "end": {
                        "dateTime": meeting.stop.strftime('%Y-%m-%d T %H:%M:%S') if meeting.stop else meeting.stop,
                        "timeZone": "UTC"
                    },
                    "showAs": meeting.show_as,
                    "location": {
                        "displayName": meeting.location if meeting.location else "",
                    },

                }
                if meeting.recurrency:
                    payload.update({"recurrence": {
                        "pattern": {
                            "daysOfWeek": self.getdays(meeting),
                            "type": (
                                        'Absolute' if meeting.rrule_type != "weekly" and meeting.rrule_type != "daily" else "") + meeting.rrule_type,
                            "interval": meeting.interval,
                            "month": int(meeting.start.month),  # (meeting.start[5] + meeting.start[6]),
                            "dayOfMonth": int(meeting.start.day),  # (meeting.start[8] + meeting.start[9]),
                            "firstDayOfWeek": "sunday",
                            # "index": "first"
                        },
                        "range": {
                            "type": "endDate",
                            "startDate": meeting.start.strftime('%Y-%m-%d'),  # meeting.start[:10],
                            "endDate": meeting.final_date,
                            "recurrenceTimeZone": "UTC",
                            "numberOfOccurrences": meeting.count,
                        }
                    }})
                # else:
                #     payload.update({"recurrence": {}})
                if meeting.name not in added:
                    response = requests.patch(
                        'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events/' + meeting.office_id,
                        headers=header, data=json.dumps(payload)).content

                    self.env.cr.commit()
                    if meeting.recurrency:
                        added.append(meeting.name)

        except Exception as e:
            raise ValidationError(_(str(e)))

        # raise osv.except_osv(_("Success!"), (_(" Sync Successfully !")))

    def getAttendee(self, attendees):
        """
        Get attendees from odoo and convert to attendees Office365 accepting
        :param attendees:
        :return: Office365 accepting attendees

        """
        attendee_list = []
        for attendee in attendees:
            attendee_list.append({
                "status": {
                    "response": 'Accepted',
                    "time": datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
                },
                "type": "required",
                "emailAddress": {
                    "address": attendee.email,
                    "name": attendee.display_name
                }
            })
        return attendee_list

    def getTime(self, alarm):
        """
        Convert ODOO time to minutes as Office365 accepts time in minutes
        :param alarm:
        :return: time in minutes
        """
        if alarm.interval == 'minutes':
            return alarm[0].duration
        elif alarm.interval == "hours":
            return alarm[0].duration * 60
        elif alarm.interval == "days":
            return alarm[0].duration * 60 * 24

    def getdays(self, meeting):
        """
        Returns days of week the event will occure
        :param meeting:
        :return: list of days
        """
        days = []
        if meeting.su:
            days.append("sunday")
        if meeting.mo:
            days.append("monday")
        if meeting.tu:
            days.append("tuesday")
        if meeting.we:
            days.append("wednesday")
        if meeting.th:
            days.append("thursday")
        if meeting.fr:
            days.append("friday")
        if meeting.sa:
            days.append("saturday")
        return days

    @api.model
    def sync_mail_scheduler(self):
        self.sync_mail()

    def sync_mail(self):
        try:
            self.env.user.send_mail_flag = False
            self.env.cr.commit()
            self.sync_inbox_mail()
            self.sync_sent_mail()
            self.env.user.send_mail_flag = True
        except Exception as e:
            self.env.user.send_mail_flag = True
            self.env.cr.commit()
            raise ValidationError(_(str(e)))
        self.env.cr.commit()
        # raise osv.except_osv(("Success!"), (" Sync Successful !"))

    def sync_inbox_mail(self):
        """
        This function syncs odoo mails to office 365 outlook inbox
        :return:
        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'value' not in json.loads((response.decode('utf-8'))).keys():
                raise osv.except_osv(("Access TOken Expired!"), (" Please Regenerate Access Token !"))
            folders = json.loads((response.decode('utf-8')))['value']
            inbox_id = [folder['id'] for folder in folders if folder['displayName'] == 'Inbox'][0]
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders/' + inbox_id + '/messages?$top=100000&$count=true',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content

            messages = json.loads((response.decode('utf-8')))['value']
            for message in messages:

                if 'from' not in message.keys() or self.env['mail.message'].search([('office_id', '=', message['id'])]):
                    continue

                if 'address' not in message.get('from').get('emailAddress'):
                    continue

                attachment_ids = self.getAttachment(message)


                self.env.cr.commit()
                from_partner = self.env['res.partner'].search(
                    [('email', "=", message['from']['emailAddress']['address'])])
                if not from_partner:
                    from_partner = self.env['res.partner'].create({
                        'email': message['from']['emailAddress']['address'],
                        'name': message['from']['emailAddress']['name'],
                    })

                recipient_partners = []
                channel_ids = []
                for recipient in message['toRecipients']:
                    if recipient['emailAddress']['address'].lower() == self.env.user.office365_email.lower() or \
                            recipient['emailAddress'][
                                'address'].lower() == self.env.user.office365_id_address.lower():
                        to = self.env.user.email
                    else:
                        to = recipient['emailAddress']['address']

                    # to_user = self.env['res.users'].search(
                    #     [('email', "=", to)])
                    # if not to_user:
                    #     to_user = self.env['res.users'].create({
                    #         'email': to,
                    #         'login': to,
                    #         'name': recipient['emailAddress']['name'],
                    #     })
                    to_partner = self.env['res.partner'].search(
                        [('email', "=", to)])
                    to_partner = to_partner[0] if to_partner else to_partner
                    from_partner = from_partner[0] if from_partner else from_partner
                    if not to_partner:
                        to_partner = self.env['res.partner'].create({
                            'email': recipient['emailAddress']['address'],
                            'name': recipient['emailAddress']['name'],
                        })
                    recipient_partners.append(to_partner.id)
                    channel_partner = self.env['mail.channel.partner'].search(
                        [('partner_id', '=', from_partner.id), ('channel_id', 'not in', [1, 2])])

                    channel_found = False
                    for channel_prtnr in channel_partner:
                        to_chanel_partner = self.env['mail.channel.partner'].search(
                            [('partner_id', '=', to_partner.id), ('channel_id', '=', channel_prtnr.channel_id.id),
                             ('channel_id', 'not in', [1, 2])])
                        if to_chanel_partner:
                            channel_found = True
                            channel_partner = channel_prtnr
                            break

                    if not channel_found:
                        channel = self.env['mail.channel'].create({
                            'name': from_partner.name + ', ' + to_partner.name,
                            'channel_type': 'chat',
                            'public': 'private'
                        })
                        channel_ids.append(channel.id)
                        from_channel_partner = self.env['mail.channel.partner'].create({
                            'partner_id': from_partner.id,
                            'channel_id': channel.id,
                            'is_pinned': True,
                        })

                        to_channel_partner = self.env['mail.channel.partner'].create({
                            'partner_id': to_partner.id,
                            'channel_id': channel.id,
                            'is_pinned': True,
                        })
                        channel_ids.append(channel.id)

                    else:
                        channel_ids.append(channel_partner[0].channel_id.id)

                self.env['mail.message'].create({
                    'subject': message['subject'],
                    'date': message['sentDateTime'],
                    'body': message['bodyPreview'],
                    'email_from': message['from']['emailAddress']['address'],
                    'channel_ids': [[6, 0, channel_ids]],
                    'partner_ids': [[6, 0, recipient_partners]],
                    'attachment_ids': [[6, 0, attachment_ids]],
                    'office_id': message['id'],
                    'model': 'res.partner',
                    'res_id': from_partner.id,
                    'author_id': from_partner.id
                })
                self.env.cr.commit()
        except Exception as e:
            self.env.user.send_mail_flag = True
            raise ValidationError(_(str(e)))

            # raise osv.except_osv(("Success!"), (" Sync Successful !"))

    def sync_sent_mail(self):
        """
            send emial from odoo to user
        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            folders = json.loads((response.decode('utf-8')))['value']
            sentbox_id = [folder['id'] for folder in folders if folder['displayName'] == 'Sent Items'][0]
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders/' + sentbox_id + '/messages?$top=100000&$count=true',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            messages = json.loads((response.decode('utf-8')))['value']
            for message in messages:

                if 'from' not in message.keys() or self.env['mail.message'].search([('office_id', '=', message['id'])]):
                    continue
                attachment_ids = self.getAttachment(message)
                if message['from']['emailAddress']['address'].lower() == self.env.user.office365_email.lower() or \
                        message['from']['emailAddress'][
                            'address'].lower() == self.env.user.office365_id_address.lower():
                    email_from = self.env.user.email
                else:
                    email_from = message['from']['emailAddress']['address']

                # from_user = self.env['res.users'].search(
                #     [('email', "=", email_from)])
                # if not from_user:
                #     from_user = self.env['res.users'].create({
                #         'email': email_from,
                #         'login': email_from,
                #         'name': message['from']['emailAddress']['name'],
                #     })
                self.env.cr.commit()
                from_partner = self.env['res.partner'].search(
                    [('email', "=", email_from)])
                if not from_partner:
                    from_partner = self.env['res.partner'].create({
                        'email': message['from']['emailAddress']['address'],
                        'name': message['from']['emailAddress']['name'],
                    })
                from_partner = from_partner[0] if from_partner else from_partner
                recipient_partners = []
                channel_ids = []
                for recipient in message['toRecipients']:
                    to_user = self.env['res.users'].search(
                        [('email', "=", recipient['emailAddress']['address'])])
                    # if not to_user:
                    #     to_user = self.env['res.users'].create({
                    #         'email': recipient['emailAddress']['address'],
                    #         'login': recipient['emailAddress']['address'],
                    #         'name': recipient['emailAddress']['name'],
                    #     })
                    self.env.cr.commit()
                    to_partner = self.env['res.partner'].search(
                        [('email', "=", recipient['emailAddress']['address'])])
                    to_partner = to_partner[0] if to_partner else to_partner

                    if not to_partner:
                        to_partner = self.env['res.partner'].create({
                            'email': recipient['emailAddress']['address'],
                            'name': recipient['emailAddress']['name'],
                        })
                    recipient_partners.append(to_partner.id)
                    channel_partner = self.env['mail.channel.partner'].search(
                        [('partner_id', '=', to_partner.id), ('channel_id', 'not in', [1, 2])])
                    channel_found = False
                    for channel_prtnr in channel_partner:
                        from_chanel_partner = self.env['mail.channel.partner'].search(
                            [('partner_id', '=', from_partner.id), ('channel_id', '=', channel_prtnr.channel_id.id),
                             ('channel_id', 'not in', [1, 2])])
                        if from_chanel_partner:
                            channel_found = True
                            channel_partner = channel_prtnr
                            break

                    if not channel_found:
                        channel = self.env['mail.channel'].create({
                            'name': to_partner.name + ', ' + from_partner.name,
                            'channel_type': 'chat',
                            'public': 'private'
                        })
                        to_channel_partner = self.env['mail.channel.partner'].create({
                            'partner_id': to_partner.id,
                            'channel_id': channel.id,
                            'is_pinned': True,
                        })
                        from_channel_partner = self.env['mail.channel.partner'].create({
                            'partner_id': from_partner.id,
                            'channel_id': channel.id,
                            'is_pinned': True,
                        })
                        channel_ids.append(channel.id)

                    else:
                        channel_ids.append(channel_partner.channel_id.id)

                self.env['mail.message'].create({
                    'subject': message['subject'],
                    'date': message['sentDateTime'],
                    'body': message['bodyPreview'],
                    'email_from': email_from,
                    'channel_ids': [[6, 0, channel_ids]],
                    'partner_ids': [[6, 0, recipient_partners]],
                    'attachment_ids': [[6, 0, attachment_ids]],
                    'office_id': message['id'],
                    'model': 'res.partner',
                    'res_id': to_partner.id,
                    'author_id': from_partner.id
                })
                # self.env.cr.commit()
                # self.env.user.send_mail_flag = True
        except Exception as e:
            raise ValidationError(_(str(e)))

    def getAttachment(self, message):
        if self.env.user.expires_in:
            expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
            expires_in = expires_in + timedelta(seconds=3600)
            nowDateTime = datetime.now()
            if nowDateTime > expires_in:
                self.generate_refresh_token()

        response = requests.get(
            'https://graph.microsoft.com/v1.0/me/messages/' + message['id'] + '/attachments/',
            headers={
                'Host': 'outlook.office.com',
                'Authorization': 'Bearer {0}'.format(self.env.user.token),
                'Accept': 'application/json',
                'X-Target-URL': 'http://outlook.office.com',
                'connection': 'keep-Alive'
            }).content
        attachments = json.loads((response.decode('utf-8')))['value']
        attachment_ids = []
        for attachment in attachments:
            if 'contentBytes' not in attachment or 'name' not in attachment:
                continue
            odoo_attachment = self.env['ir.attachment'].create({
                'datas': attachment['contentBytes'],
                'name': attachment["name"],
                'datas_fname': attachment["name"]})
            self.env.cr.commit()
            attachment_ids.append(odoo_attachment.id)
        return attachment_ids

    @api.model
    def auto_import_tasks(self):
        self.import_tasks()

    @api.model
    def auto_export_tasks(self):
        self.export_tasks()

    def import_tasks(self):

        """
        import tast from office 365 to odoo

        :return: None
        """

        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/beta/me/outlook/tasks',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Content-type': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'value' not in json.loads((response.decode('utf-8'))).keys():
                raise osv.except_osv(response)
            tasks = json.loads((response.decode('utf-8')))['value']
            partner_model = self.env['ir.model'].search([('model', '=', 'res.partner')])
            partner = self.env['res.partner'].search([('email', '=', self.env.user.email)])
            activity_type = self.env['mail.activity.type'].search([('name', '=', 'Todo')])
            if partner_model:
                self.env.user.is_task_sync_on = True
                self.env.cr.commit()
                for task in tasks:
                    if not self.env['mail.activity'].search([('office_id', '=', task['id'])]) and task[
                        'status'] != 'completed':
                        if 'dueDateTime' in task:
                            if task['dueDateTime'] is None:
                                continue
                        else:
                            continue

                        self.env['mail.activity'].create({
                            'res_id': partner[0].id,
                            'activity_type_id': activity_type.id,
                            'summary': task['subject'],
                            'date_deadline': (
                                datetime.strptime(task['dueDateTime']['dateTime'][:-16], '%Y-%m-%dT')).strftime(
                                '%Y-%m-%d'),
                            'note': task['body']['content'],
                            'res_model_id': partner_model.id,
                            'office_id': task['id'],
                        })
                    elif self.env['mail.activity'].search([('office_id', '=', task['id'])]) and task[
                        'status'] != 'completed':
                        activity = self.env['mail.activity'].search([('office_id', '=', task['id'])])[0]
                        activity.write({
                            'res_id': partner[0].id,
                            'activity_type_id': activity_type.id,
                            'summary': task['subject'],
                            'date_deadline': (
                                datetime.strptime(task['dueDateTime']['dateTime'][:-16], '%Y-%m-%dT')).strftime(
                                '%Y-%m-%d'),
                            'note': task['body']['content'],
                            'res_model_id': partner_model.id,
                            'office_id': task['id'],
                        })
                    elif self.env['mail.activity'].search([('office_id', '=', task['id'])]) and task[
                        'status'] == 'completed':
                        activity = self.env['mail.activity'].search([('office_id', '=', task['id'])])[0]
                        activity.unlink()

                    self.env.cr.commit()

            odoo_activities = self.env['mail.activity'].search(
                [('office_id', '!=', None), ('res_id', '=', self.env.user.partner_id.id)])
            task_ids = [task['id'] for task in tasks]
            for odoo_activity in odoo_activities:
                if odoo_activity.office_id not in task_ids:
                    odoo_activity.unlink()
                    self.env.cr.commit()
            self.env.user.is_task_sync_on = False
            self.env.cr.commit()

        except Exception as e:
            self.env.user.is_task_sync_on = False
            self.env.cr.commit()
            raise ValidationError(_(str(e)))
        # raise osv.except_osv(_("Success!"), (_(" Tasks are  Successfully Imported !")))

    def export_tasks(self):
        if self.env.user.expires_in:
            expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
            expires_in = expires_in + timedelta(seconds=3600)
            nowDateTime = datetime.now()
            if nowDateTime > expires_in:
                self.generate_refresh_token()

        odoo_activities = self.env['mail.activity'].search([('res_id', '=', self.env.user.partner_id.id)])
        for activity in odoo_activities:
            url = 'https://graph.microsoft.com/beta/me/outlook/tasks'
            if activity.office_id:
                url += '/' + activity.office_id

            data = {
                'subject': activity.summary if activity.summary else activity.note,
                "body": {
                    "contentType": "html",
                    "content": activity.note
                },
                "dueDateTime": {
                    "dateTime": str(activity.date_deadline) + 'T00:00:00Z',
                    "timeZone": "UTC"
                },
            }
            if activity.office_id:

                response = requests.patch(
                    url, data=json.dumps(data),
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.env.user.token),
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content
            else:
                response = requests.post(
                    url, data=json.dumps(data),
                    headers={

                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content

                if 'id' not in json.loads((response.decode('utf-8'))).keys():
                    raise osv.except_osv(_("Error!"), (_(response["error"])))
                activity.office_id = json.loads((response.decode('utf-8')))['id']
            self.env.cr.commit()

            raise osv.except_osv(_("Success!"), (_("Tasks are Successfully exported! !")))

    def developer_test(self):
        try:
            channel = self.env['mail.channel'].search()
            raise osv.except_osv(_("Error!"), (_(channel)))
        except Exception as e:
            # self.env.user.send_mail_flag = True
            self.env.cr.commit()
            raise ValidationError(_(str(e)))
        self.env.cr.commit()

    @api.model
    def sync_customer_mail_scheduler(self):
        self.sync_customer_mail()

    def sync_customer_mail(self):
        try:
            # self.env.user.send_mail_flag = False
            self.env.cr.commit()
            self.sync_customer_inbox_mail()
            self.sync_customer_sent_mail()
            # self.env.user.send_mail_flag = True
        except Exception as e:
            # self.env.user.send_mail_flag = True
            self.env.cr.commit()
            raise ValidationError(_(str(e)))

        self.env.cr.commit()

        raise osv.except_osv(("Success!"), ("Mails Synced Successfully !"))

    def sync_customer_inbox_mail(self):
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'value' not in json.loads((response.decode('utf-8'))).keys():
                raise osv.except_osv("Access TOken Expired!", " Please Regenerate Access Token !")
            folders = json.loads((response.decode('utf-8')))['value']
            inbox_id = [folder['id'] for folder in folders if folder['displayName'] == _('Inbox')][0]
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders/' + inbox_id + '/messages?$top=100000&$count=true',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content

            messages = json.loads((response.decode('utf-8')))['value']
            for message in messages:
                if 'from' not in message.keys() or self.env['mail.message'].search([('office_id', '=', message['id'])]):
                    continue

                if 'address' not in message.get('from').get('emailAddress') or message['bodyPreview'] == "":
                    continue

                attachment_ids = self.getAttachment(message)

                from_partner = self.env['res.partner'].search(
                    [('email', "=", message['from']['emailAddress']['address'])])
                if not from_partner:
                    continue
                from_partner = from_partner[0] if from_partner else from_partner
                # if from_partner:
                #     from_partner = from_partner[0]
                recipient_partners = []
                channel_ids = []
                for recipient in message['toRecipients']:
                    if recipient['emailAddress']['address'].lower() == self.env.user.office365_email.lower() or \
                            recipient['emailAddress'][
                                'address'].lower() == self.env.user.office365_id_address.lower():
                        to_user = self.env['res.users'].search(
                            [('id', "=", self._uid)])
                    else:
                        to = recipient['emailAddress']['address']
                        to_user = self.env['res.users'].search(
                            [('office365_id_address', "=", to)])
                        to_user = to_user[0] if to_user else to_user

                    if to_user:
                        to_partner = to_user.partner_id
                        recipient_partners.append(to_partner.id)

                self.env['mail.message'].create({
                    'subject': message['subject'],
                    'date': message['sentDateTime'],
                    'body': message['bodyPreview'],
                    'email_from': message['from']['emailAddress']['address'],
                    # 'channel_ids': [[6, 0, channel_ids]],
                    'partner_ids': [[6, 0, recipient_partners]],
                    'attachment_ids': [[6, 0, attachment_ids]],
                    'office_id': message['id'],
                    'author_id': from_partner.id,
                    'model': 'res.partner',
                    'res_id': from_partner.id
                })
                self.env.cr.commit()
        except Exception as e:
            # self.env.user.send_mail_flag = True
            raise ValidationError(_(str(e)))

        raise osv.except_osv(("Success!"), (" Sync Successful !"))

    def sync_customer_sent_mail(self):
        """
        :return:
        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            folders = json.loads((response.decode('utf-8')))['value']
            sentbox_id = [folder['id'] for folder in folders if folder['displayName'] == _('Sent Items')][0]
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders/' + sentbox_id + '/messages?$top=100000&$count=true',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            messages = json.loads((response.decode('utf-8')))['value']
            for message in messages:

                if 'from' not in message.keys() or self.env['mail.message'].search([('office_id', '=', message['id'])]):
                    continue

                if message['bodyPreview'] == "":
                    continue

                attachment_ids = self.getAttachment(message)
                if message['from']['emailAddress']['address'].lower() == self.env.user.office365_email.lower() or \
                        message['from']['emailAddress'][
                            'address'].lower() == self.env.user.office365_id_address.lower():
                    email_from = self.env.user.email
                else:
                    email_from = message['from']['emailAddress']['address']

                from_user = self.env['res.users'].search(
                    [('id', "=", self._uid)])
                if from_user:
                    from_partner = from_user.partner_id
                else:
                    continue

                channel_ids = []
                for recipient in message['toRecipients']:

                    to_partner = self.env['res.partner'].search(
                        [('email', "=", recipient['emailAddress']['address'])])
                    to_partner = to_partner[0] if to_partner else to_partner

                    if not to_partner:
                        continue


                    self.env['mail.message'].create({
                        'subject': message['subject'],
                        'date': message['sentDateTime'],
                        'body': message['bodyPreview'],
                        'email_from': email_from,
                        # 'channel_ids': [[6, 0, channel_ids]],
                        'partner_ids': [[6, 0, [to_partner.id]]],
                        'attachment_ids': [[6, 0, attachment_ids]],
                        'office_id': message['id'],
                        'author_id': from_partner.id,
                        'model': 'res.partner',
                        'res_id': to_partner.id
                    })
                    self.env.cr.commit()
                # self.env.user.send_mail_flag = True
        except Exception as e:
            raise ValidationError(_(str(e)))
        """
        :return:
        """
        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            folders = json.loads((response.decode('utf-8')))['value']
            sentbox_id = [folder['id'] for folder in folders if folder['displayName'] == _('Sent Items')][0]
            response = requests.get(
                'https://graph.microsoft.com/v1.0/me/mailFolders/' + sentbox_id + '/messages?$top=100000&$count=true',
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            messages = json.loads((response.decode('utf-8')))['value']
            for message in messages:

                if 'from' not in message.keys() or self.env['mail.message'].search([('office_id', '=', message['id'])]):
                    continue

                if message['bodyPreview'] == "":
                    continue

                attachment_ids = self.getAttachment(message)
                if message['from']['emailAddress']['address'].lower() == self.env.user.office365_email.lower() or \
                        message['from']['emailAddress'][
                            'address'].lower() == self.env.user.office365_id_address.lower():
                    email_from = self.env.user.email
                else:
                    email_from = message['from']['emailAddress']['address']

                from_user = self.env['res.users'].search(
                    [('id', "=", self._uid)])
                if from_user:
                    from_partner = from_user.partner_id
                else:
                    continue

                channel_ids = []
                for recipient in message['toRecipients']:

                    to_partner = self.env['res.partner'].search(
                        [('email', "=", recipient['emailAddress']['address'])])
                    to_partner = to_partner[0] if to_partner else to_partner

                    if not to_partner:
                        to_partner = self.env['res.partner'].create({
                            'email': recipient['emailAddress']['address'],
                            'name': recipient['emailAddress']['name'],
                        })


                    self.env['mail.message'].create({
                        'subject': message['subject'],
                        'date': message['sentDateTime'],
                        'body': message['bodyPreview'],
                        'email_from': email_from,
                        # 'channel_ids': [[6, 0, channel_ids]],
                        'partner_ids': [[6, 0, [to_partner.id]]],
                        'attachment_ids': [[6, 0, attachment_ids]],
                        'office_id': message['id'],
                        'author_id': from_partner.id,
                        'model': 'res.partner',
                        'res_id': to_partner.id
                    })
                    self.env.cr.commit()
                # self.env.user.send_mail_flag = True
        except Exception as e:
            raise ValidationError(_(str(e)))

    def generate_refresh_token(self):
        header = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            data='grant_type=refresh_token&refresh_token=' + self.env.user.refresh_token + '&redirect_uri=' + self.env.user.redirect_url + '&client_id=' + self.env.user.client_id + '&client_secret=' + self.env.user.secret
            , headers=header).content

        response = json.loads((str(response)[2:])[:-1])
        if 'access_token' not in response:
            response["error_description"] = response["error_description"].replace("\\r\\n", " ")
            raise osv.except_osv(_("Error!"), (_(response["error"] + " " + response["error_description"])))
        else:
            self.env.user.token = response['access_token']
            self.env.user.refresh_token = response['refresh_token']
            self.env.user.expires_in = int(round(time.time() * 1000))

    def export_contacts(self):

        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            odoo_contacts = self.env['res.partner'].search([])
            url =  'https://graph.microsoft.com/v1.0/me/contacts'

            headers = {

                'Host': 'outlook.office365.com',
                'Authorization': 'Bearer {0}'.format(self.env.user.token),
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'X-Target-URL': 'http://outlook.office.com',
                'connection': 'keep-Alive'

            }

            response = requests.get(
                url, headers=headers
            ).content
            response = json.loads(response.decode('utf-8'))
            office_contact = [response['value'][i]['emailAddresses'][0]['address'] for i in range(len(response['value']))]


            for contact in odoo_contacts:
                data = {
                    "givenName": contact.name,
                    "emailAddresses": [
                        {
                            "address": contact.email
                        }
                    ],
                }
                if contact.email in office_contact:
                    continue

                post_response = requests.post(
                    url,data=json.dumps(data), headers=headers
                ).content

                if 'id' not in json.loads(post_response.decode('utf-8')).keys():
                    raise osv.except_osv(_("Error!"), (_(post_response["error"])))

        except Exception as e:
            raise ValidationError(_(str(e)))

        raise osv.except_osv(_("Success!"), (_("Contacts are Successfully exported!!")))

        try:
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            odoo_contacts = self.env['res.partner'].search([])
            url = 'https://graph.microsoft.com/v1.0/me/contacts'
            for contact in odoo_contacts:
                data = {
                    "givenName": contact.name,
                    "emailAddresses": [
                        {
                            "address": contact.email,
                        }
                    ],
                }
                response = requests.post(
                    url, data=json.dumps(data),
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content

                if 'id' not in json.loads((response.decode('utf-8'))).keys():
                    raise osv.except_osv(_("Error!"), (_(response["error"])))

                raise osv.except_osv(_("Success!"), (_("Tasks are Successfully exported! !")))
        except Exception as e:
            raise ValidationError(_(str(e)))


class CustomMeeting(models.Model):
    """
    adding office365 event ID to ODOO meeting to remove duplication and facilitate updation
    """
    _inherit = 'calendar.event'

    office_id = fields.Char('Office365 Id')


class CustomMessage(models.Model):
    """
    Email will be sent to the recipient of the message.
    """
    _inherit = 'mail.message'

    office_id = fields.Char('Office Id')

    @api.model
    def create(self, values):
        """
        overriding create message to send email on message creation
        :param values:
        :return:
        """
        ################## New Code ##################
        o365_id = None

        if self.env.user.send_mail_flag and self.env.user.token and 'res_id' in values.keys():
            if self.env.user.expires_in:
                expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

        if "subject" in values:
            # attachments_list = self.getAttachments(values['attachment_ids'])
            subject = values['subject'] if values['subject'] else values['body']
            body = values['body']
            to_partner = None
            if values['model'] == 'mail.channel':
                from_partner = self.env['res.partner'].search([('email', '=', self.env.user.email)])
                from_partner = from_partner[0] if from_partner else from_partner
                channel_partner = self.env['mail.channel.partner'].search(
                    [('channel_id', '=', values['res_id']), ('partner_id', '!=', from_partner.id)])
                to_partner = channel_partner[0].partner_id if channel_partner else channel_partner
            if values['model'] == 'account.invoice' or values['model'] == 'crm.lead' or values[
                'model'] == 'sale.order' or values['model'] == 'res.partner':
                obj = self.env[values['model']].search([('id', '=', values['res_id'])])
                if values['model'] == 'res.partner':
                    to_partner = obj
                else:
                    to_partner = obj.partner_id

            if to_partner and "office_id" not in values:
                if not self.env.user.token or not body or body == "" or body == (_("Contact created")) or body == _(
                        "Quotation created") or _("has joined the My Company network.") in body:
                    not_send_email = ""
                elif to_partner.email == False:
                    raise osv.except_osv(_("Error!"),
                                         (_("Unable to send email, please fill the Contact's email address.")))
                else:
                    to_user = self.env['res.users'].search([('email', '=', to_partner.email)])
                    email = ""
                    if to_user and to_user.office365_email:
                        email = to_user.office365_email
                    else:
                        email = to_partner.email
                    data = {
                        "message": {
                            "subject": subject if subject else body,

                            "body": {
                                "contentType": "Text",
                                "content": body
                            },
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        'address': email
                                    }
                                }
                            ],
                            "ccRecipients": []
                        },
                        "saveToSentItems": "true"
                    }
                    new_data = {
                        "subject": subject,
                        "importance": "Low",
                        "body": {
                            "contentType": "HTML",
                            "content": body
                        },
                        "toRecipients": [
                            {
                                "emailAddress": {
                                    "address": email
                                }
                            }
                        ]
                    }

                    response = requests.post(
                        'https://graph.microsoft.com/v1.0/me/messages', data=json.dumps(new_data),
                        headers={
                            'Host': 'outlook.office.com',
                            'Authorization': 'Bearer {0}'.format(self.env.user.token),
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-Target-URL': 'http://outlook.office.com',
                            'connection': 'keep-Alive'
                        })
                    if 'id' in json.loads((response.content.decode('utf-8'))).keys():
                        o365_id = json.loads((response.content.decode('utf-8')))['id']
                        if values['attachment_ids']:
                            for attachment in self.getAttachments(values['attachment_ids']):
                                attachment_response = requests.post(
                                    'https://graph.microsoft.com/beta/me/messages/' + o365_id + '/attachments',
                                    data=json.dumps(attachment),
                                    headers={
                                        'Host': 'outlook.office.com',
                                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                                        'Accept': 'application/json',
                                        'Content-Type': 'application/json',
                                        'X-Target-URL': 'http://outlook.office.com',
                                        'connection': 'keep-Alive'
                                    })
                        send_response = requests.post(
                            'https://graph.microsoft.com/v1.0/me/messages/' + o365_id + '/send',
                            headers={
                                'Host': 'outlook.office.com',
                                'Authorization': 'Bearer {0}'.format(self.env.user.token),
                                'Accept': 'application/json',
                                'Content-Type': 'application/json',
                                'X-Target-URL': 'http://outlook.office.com',
                                'connection': 'keep-Alive'
                            })
                        if send_response.status_code == 401:
                            raise osv.except_osv(("Access Token Expired!"), (" Please Regenerate Access Token !"))

                        if send_response.status_code != 202:
                            raise osv.except_osv(_("Error!"), (_(
                                "Mail not sent to " + to_partner.email + ' with status code ' + str(
                                    response.status_code))))

                    if response.status_code == 401:
                        raise osv.except_osv(_("Access Token Expired!"), (_(
                            "Please Regenerate Access Token or go to office365 settings in users and in Mails tab disable checkbox")))
                    elif response.status_code != 201:
                        raise osv.except_osv(_("Error!"), (_(
                            "Mail not sent to " + to_partner.email + ' with status code ' + str(response.status_code))))
            else:
                if "office_id" in values:
                    o365_id = values['office_id']
        ################## New Code ##################
        # coming from mail.js that does not have pid in its values

        if self.env.context.get('default_starred'):
            self = self.with_context({'default_starred_partner_ids': [(4, self.env.user.partner_id.id)]})

        if 'email_from' not in values:  # needed to compute reply_to
            values['email_from'] = self._get_default_from()
        if not values.get('message_id'):
            values['message_id'] = self._get_message_id(values)
        if 'reply_to' not in values:
            values['reply_to'] = self._get_reply_to(values)
        if 'record_name' not in values and 'default_record_name' not in self.env.context:
            values['record_name'] = self._get_record_name(values)

        if 'attachment_ids' not in values:
            values.setdefault('attachment_ids', [])

        # extract base64 images
        if 'body' in values:
            Attachments = self.env['ir.attachment']
            data_to_url = {}

            def base64_to_boundary(match):
                key = match.group(2)
                if not data_to_url.get(key):
                    name = 'image%s' % len(data_to_url)
                    attachment = Attachments.create({
                        'name': name,
                        'datas': match.group(2),
                        'datas_fname': name,
                        'res_model': 'mail.message',
                    })
                    values['attachment_ids'].append((4, attachment.id))
                    data_to_url[key] = '/web/image/%s' % attachment.id
                return '%s%s alt="%s"' % (data_to_url[key], match.group(3), name)

            values['body'] = _image_dataurl.sub(base64_to_boundary, tools.ustr(values['body']))

        message = super(CustomMessage, self).create(values)
        message._invalidate_documents()

        # if not self.env.context.get('message_create_from_mail_mail'):
        #     message._notify(force_send=self.env.context.get('mail_notify_force_send', True),
        #                     user_signature=self.env.context.get('mail_notify_user_signature', True))
        if o365_id:
            message.office_id = o365_id

        return message

    def getAttachments(self, attachment_ids):
        attachment_list = []
        if attachment_ids:
            attachments = self.env['ir.attachment'].search([('id', 'in', [id[1] for id in attachment_ids])])
            for attachment in attachments:
                attachment_list.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.name,
                    "contentBytes": attachment.datas.decode("utf-8")
                })
        return attachment_list

    def generate_refresh_token(self):
        header = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            data='grant_type=refresh_token&refresh_token=' + self.env.user.refresh_token + '&redirect_uri=' + self.env.user.redirect_url + '&client_id=' + self.env.user.client_id + '&client_secret=' + self.env.user.secret
            , headers=header).content

        response = json.loads((str(response)[2:])[:-1])
        if 'access_token' not in response:
            response["error_description"] = response["error_description"].replace("\\r\\n", " ")
            raise osv.except_osv(_("Error!"), (_(response["error"] + " " + response["error_description"])))
        else:
            self.env.user.token = response['access_token']
            self.env.user.refresh_token = response['refresh_token']
            self.env.user.expires_in = int(round(time.time() * 1000))


class CustomActivity(models.Model):
    _inherit = 'mail.activity'

    office_id = fields.Char('Office365 Id')

    @api.model
    def create(self, values):
        if self.env.user.expires_in:
            expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
            expires_in = expires_in + timedelta(seconds=3600)
            nowDateTime = datetime.now()
            if nowDateTime > expires_in:
                self.generate_refresh_token()

        o365_id = None
        if self.env.user.office365_email and not self.env.user.is_task_sync_on and values[
            'res_id'] == self.env.user.partner_id.id:
            data = {
                'subject': values['summary'] if values['summary'] else values['note'],
                "body": {
                    "contentType": "html",
                    "content": values['note']
                },
                "dueDateTime": {
                    "dateTime": values['date_deadline'] + 'T00:00:00Z',
                    "timeZone": "UTC"
                },
            }
            response = requests.post(
                'https://graph.microsoft.com/beta/me/outlook/tasks', data=json.dumps(data),
                headers={
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(self.env.user.token),
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }).content
            if 'id' in json.loads((response.decode('utf-8'))).keys():
                o365_id = json.loads((response.decode('utf-8')))['id']

        """
        original code!
        """

        activity = super(CustomActivity, self).create(values)
        self.env[activity.res_model].browse(activity.res_id).message_subscribe(
            partner_ids=[activity.user_id.partner_id.id])
        if activity.date_deadline <= fields.Date.today():
            self.env['bus.bus'].sendone(
                (self._cr.dbname, 'res.partner', activity.user_id.partner_id.id),
                {'type': 'activity_updated', 'activity_created': True})
        if o365_id:
            activity.office_id = o365_id
        return activity

    def generate_refresh_token(self):
        header = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            data='grant_type=refresh_token&refresh_token=' + self.env.user.refresh_token + '&redirect_uri=' + self.env.user.redirect_url + '&client_id=' + self.env.user.client_id + '&client_secret=' + self.env.user.secret
            , headers=header).content

        response = json.loads((str(response)[2:])[:-1])
        if 'access_token' not in response:
            response["error_description"] = response["error_description"].replace("\\r\\n", " ")
            raise osv.except_osv(("Error!"), (response["error"] + " " + response["error_description"]))
        else:
            self.env.user.token = response['access_token']
            self.env.user.refresh_token = response['refresh_token']
            self.env.user.expires_in = int(round(time.time() * 1000))

    @api.multi
    def unlink(self):
        for activity in self:
            if activity.office_id:
                response = requests.delete(
                    'https://graph.microsoft.com/beta/me/outlook/tasks/' + activity.office_id,
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    })
                if response.status_code != 204 and response.status_code != 404:
                    raise osv.except_osv(_("Office365 SYNC ERROR"), (_("Error: " + str(response.status_code))))
            if activity.date_deadline <= fields.Date.today():
                self.env['bus.bus'].sendone(
                    (self._cr.dbname, 'res.partner', activity.user_id.partner_id.id),
                    {'type': 'activity_updated', 'activity_deleted': True})
        return super(CustomActivity, self).unlink()


class Office365Credentials(models.Model):

    _name = 'office.credentials'

    login_url = fields.Char('Login URL', compute='_compute_url', readonly=True)
    code = fields.Char('code')
    field_name = fields.Char('office')
    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret = fields.Char('Secret')
    microsoft_email = fields.Char(String='Microsoft Account Email')
    microsoft_password = fields.Char(String='Microsoft Account Password')