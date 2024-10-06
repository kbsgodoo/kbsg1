# -*- coding: utf-8 -*-

import logging
import re

from odoo import fields, models, api, osv
from odoo.exceptions import ValidationError

from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools
from odoo.exceptions import UserError, AccessError
import requests
import json
from datetime import datetime
import time
from datetime import timedelta

_logger = logging.getLogger(__name__)
_image_dataurl = re.compile(r'(data:image/[a-z]+?);base64,([a-z0-9+/]{3,}=*)([\'"])', re.I)

class CustomMeeting(models.Model):
    _inherit = 'calendar.event'

    office_id = fields.Char('Office365 Id')
    category_name = fields.Char('Categories', )
    is_update = fields.Boolean('Is Updated')
    modified_date = fields.Datetime('Modified Date')
    calendar_id = fields.Many2one(comodel_name="office.calendars", string="Office 365 Calendar Type", required=False, )
    show_as = fields.Selection(selection=[('free', 'Free'), ('busy', 'Busy'),
                                          ('tentative', 'Tentative'), ('workingElsewhere', 'Working Elsewhere'),
                                          ('oof', 'Out of Office')])

    def unlink(self):
        events = self
        res = False
        for event in events:
            if event.office_id and event.env.user.event_del_flag:
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
                response = requests.delete(
                    'https://graph.microsoft.com/v1.0/me/events/' + event.office_id,
                    headers=header)
                if response.status_code == 204:
                    if event.recurrency:
                        events_dels = self.env['calendar.event'].search([('recurrence_id', '=', event.recurrence_id.id)])
                        for event_del in events_dels:
                            event_del.write({
                                'office_id': None
                            })
                            self.env.cr.commit()
                    _logger.info('successfull deleted event ' + event.name + "from Office365 Calendar")

            res = super(CustomMeeting, event).unlink()
        return res

    def generate_refresh_token(self):
        try:
            odooUser = self.env.user
            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
            payload = 'grant_type=refresh_token&refresh_token=' + odooUser.refresh_token + \
                      '&redirect_uri=' + odooUser.redirect_url + '&client_id=' + odooUser.client_id + \
                      '&client_secret=' + odooUser.client_secret
            response = requests.post(url, data=payload, headers=header)

            if response.status_code == 200:
                newToken = json.loads((response.content.decode('utf-8')))
                odooUser.write({
                    'token': newToken['access_token'],
                    'refresh_token': newToken['refresh_token'],
                    'expires_in': int(round(time.time() * 1000))
                })
                self.env.cr.commit()
            else:
                raise ValidationError(str("Your Token has been expired and while updating automatically we are facing issue can you please "
                                          "check your credentials or again login with your Office 365 account"))
        except Exception as e:
            raise ValidationError(str(e))

class CustomMessageInbox(models.Model):
    _inherit = 'mail.message'

    office_id = fields.Char('Office Id')

class CustomMessage(models.Model):
    _inherit = 'mail.mail'

    office_id = fields.Char('Office Id')

    def create(self, values):
        """
        overriding create message to send email on message creation
        :param values:
        :return:
        """
        ################## New Code ##################
        ################## New Code ##################
        o365_id = None
        conv_id = None
        user = self.env.user
        if user.token:
            if user.expires_in:
                expires_in = datetime.fromtimestamp(int(user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()

            if 'mail_message_id' in values:
                email_obj = self.env['mail.message'].search([('id', '=', values['mail_message_id'])])
                partner_id = values['recipient_ids'][0][1]
                partner_obj = self.env['res.partner'].search([('id', '=', partner_id)])

                new_data = {
                            "subject": values['subject'] if values['subject'] else email_obj.body,
                            # "importance": "high",
                            "body": {
                                "contentType": "HTML",
                                "content": email_obj.body
                            },
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        "address": partner_obj.email
                                    }
                                }
                            ]
                        }

                response = requests.post(
                    'https://graph.microsoft.com/v1.0/me/messages', data=json.dumps(new_data),
                                        headers={
                                            'Host': 'outlook.office.com',
                                            'Authorization': 'Bearer {0}'.format(user.token),
                                            'Accept': 'application/json',
                                            'Content-Type': 'application/json',
                                            'X-Target-URL': 'http://outlook.office.com',
                                            'connection': 'keep-Alive'
                                        })
                if 'conversationId' in json.loads((response.content.decode('utf-8'))).keys():
                    conv_id = json.loads((response.content.decode('utf-8')))['conversationId']

                if 'id' in json.loads((response.content.decode('utf-8'))).keys():

                    o365_id = json.loads((response.content.decode('utf-8')))['id']
                    if email_obj.attachment_ids:
                        for attachment in self.getAttachments(email_obj.attachment_ids):
                            attachment_response = requests.post(
                                'https://graph.microsoft.com/beta/me/messages/' + o365_id + '/attachments',
                                data=json.dumps(attachment),
                                headers={
                                    'Host': 'outlook.office.com',
                                    'Authorization': 'Bearer {0}'.format(user.token),
                                    'Accept': 'application/json',
                                    'Content-Type': 'application/json',
                                    'X-Target-URL': 'http://outlook.office.com',
                                    'connection': 'keep-Alive'
                                })
                    send_response = requests.post(
                        'https://graph.microsoft.com/v1.0/me/messages/' + o365_id + '/send',
                        headers={
                            'Host': 'outlook.office.com',
                            'Authorization': 'Bearer {0}'.format(user.token),
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-Target-URL': 'http://outlook.office.com',
                            'connection': 'keep-Alive'
                        })

                    message = super(CustomMessage, self).create(values)
                    message.email_from = None

                    if conv_id:
                        message.office_id = conv_id

                    return message
                else:
                    pass
                    # print('Check your credentials! Mail does not send due to invlide office365 credentials ')
            else:
                return super(CustomMessage, self).create(values)

        else:
            # print('Office354 Token is missing.. Please add your account token and try again!')
            return super(CustomMessage, self).create(values)

    def getAttachments(self, attachment_ids):
        attachment_list = []
        if attachment_ids:
            # attachments = self.env['ir.attachment'].browse([id[0] for id in attachment_ids])
            attachments = self.env['ir.attachment'].search([('id', 'in', [i.id for i in attachment_ids])])
            for attachment in attachments:
                attachment_list.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.name,
                    "contentBytes": attachment.datas.decode("utf-8")
                })
        return attachment_list

    def generate_refresh_token(self):
        try:
            odooUser = self.env.user
            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
            payload = 'grant_type=refresh_token&refresh_token=' + odooUser.refresh_token + \
                      '&redirect_uri=' + odooUser.redirect_url + '&client_id=' + odooUser.client_id + \
                      '&client_secret=' + odooUser.client_secret
            response = requests.post(url, data=payload, headers=header)

            if response.status_code == 200:
                newToken = json.loads((response.content.decode('utf-8')))
                odooUser.write({
                    'token': newToken['access_token'],
                    'refresh_token': newToken['refresh_token'],
                    'expires_in': int(round(time.time() * 1000))
                })
                self.env.cr.commit()
            else:
                raise ValidationError(str("Your Token has been expired and while updating automatically we are facing issue can you please "
                                          "check your credentials or again login with your Office 365 account"))
        except Exception as e:
            raise ValidationError(str(e))

class CustomActivity(models.Model):
    _inherit = 'mail.activity'

    office_id = fields.Char('Office365 Id')
    is_update = fields.Boolean('Is Updated')
    modified_date = fields.Datetime('Modified Date')

class CustomContacts(models.Model):
    _inherit = 'res.partner'

    office_contact_id = fields.Char('Office365 Id')
    modified_date = fields.Datetime('Modified Date')
    location = fields.Char(string='Location')
    firstName = fields.Char(string='First Name')
    lastName = fields.Char(string='Last Name')
    middleName = fields.Char(string='Middle Name')
    is_update = fields.Boolean('Is Updated')

class ContactCateg(models.Model):
    _inherit = 'res.partner.category'

    categ_id = fields.Char(string="o_category id", required=False, )

class CalendarEventCateg(models.Model):
    _inherit = 'calendar.event.type'

    color = fields.Char(string="Color", required=False, )
    categ_id = fields.Char(string="o_category id", required=False, )

class Office365UserModel(models.Model):
    _inherit = 'res.users'

    redirect_url = fields.Char(string='Redirect URL')
    client_id = fields.Char(string='Client ID', config_parameter='odoo_office365.client_id')
    client_secret = fields.Char(string='Client Secret', config_parameter='odoo_office365.client_secret')
    login_url = fields.Char('Login URL', compute='_compute_url', config_parameter='odoo_office365.login_url')
    code = fields.Char('code')
    token = fields.Char('Token', readonly=True)
    refresh_token = fields.Char('Refresh Token', readonly=True)
    expires_in = fields.Char('Expires IN', readonly=True)
    event_del_flag = fields.Boolean('Delete events from Office365 calendar when delete in Odoo.',groups="base.group_user")
    office365_event_del_flag = fields.Boolean('Delete event from Odoo, if the event is deleted from Office 365.',groups="base.group_user")
    webhook_redirect_url = fields.Char('Webhook Redirect Url')


    @api.depends('redirect_url', 'client_id', 'client_secret')
    def _compute_url(self):
        for i in self:
            i.login_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' \
                             'client_id=%s&redirect_uri=%s&response_type=code&scope=openid+offline_access+' \
                             'Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+User.ReadWrite+Tasks.ReadWrite+' \
                             'Contacts.ReadWrite+MailboxSettings.Read' % (
                                 i.client_id, i.redirect_url)

    def get_log(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': 'Office 365 Logs',
            'view_mode': 'tree',
            'res_model': 'sync.history',
            'context': "{'create': False}"
        }

    def generate_token(self, code):
        try:
            if not self.client_id or not self.redirect_url or not self.client_secret:
                ValidationError("Please ask Admin to add Office 365 credentials")

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
            payload = 'grant_type=authorization_code&code=' + code + '&redirect_uri=' + self.redirect_url + '&client_id=' + self.client_id + '&client_secret=' + self.client_secret
            response = requests.post(url, data=payload, headers=header)
            data = {}
            if response.status_code == 200:
                response = json.loads(response.content)
                data['token'] = response['access_token']
                data['refresh_token'] = response['refresh_token']
                data['expires_in'] = response['expires_in']

                categoriesUrl = 'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'
                newDataHeader = {
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(data['token']),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }
                categoriesResponse = requests.get(categoriesUrl, headers=newDataHeader)
                if categoriesResponse.status_code == 200:
                    categories = json.loads(categoriesResponse.content)
                    self.createPartnerCategory(categories)
                    self.createCalendarCategory(categories)
                calendarUrl = 'https://graph.microsoft.com/v1.0/me/calendars'
                calendarsResponse = requests.get(calendarUrl, headers=newDataHeader)
                if calendarsResponse.status_code == 200:
                    calendars = json.loads(calendarsResponse.content)
                    self.get_calendars(calendars)
                return data
            else:
                data['error'] = 'Something went wrong while fetching token from Office 365.'
                return data

        except Exception as e:
            raise ValidationError('Invalid Credentials!')

    def createPartnerCategory(self, categories):
        for category in categories['value']:
            odooCategory = self.env['res.partner.category'].search(['|', ('categ_id', '=', category['id']), ('name', '=', category['displayName'])])
            if odooCategory:
                odooCategory.write({
                    'categ_id': category['id'],
                    'name': category['displayName'],
                })
            else:
                self.env['res.partner.category'].create({
                    'categ_id': category['id'],
                    'name': category['displayName'],
                })

    def createCalendarCategory(self, categories):
        for category in categories['value']:
            odooCategory = self.env['calendar.event.type'].search(['|', ('categ_id', '=', category['id']), ('name', '=', category['displayName'])])
            if odooCategory:
                odooCategory.write({
                    'categ_id': category['id'],
                    'color': category['color'],
                    'name': category['displayName'],
                })
            else:
                self.env['calendar.event.type'].create({
                    'categ_id': category['id'],
                    'color': category['color'],
                    'name': category['displayName'],
                })

    def get_calendars(self, calendars):
        for calendar in calendars['value']:
            odooCalendar = self.env['office.calendars'].search([('calendar_id', '=', calendar['id'])])
            if odooCalendar:
                odooCalendar.write({
                    'calendar_id': calendar['id'],
                    'name': calendar['name'],
                })
            else:
                self.env['office.calendars'].create({
                    'calendar_id': calendar['id'],
                    'name': calendar['name'],
                })

class OfficeWebhookType(models.Model):
    _name = "office.webhook"
    _description = "office365 Webhoook"

    name = fields.Char(string="Webhook Name")
    subscription_id= fields.Char(string="Subscription Id")
    office_webhook_change_type = fields.Char(string="Webhook Change Type")
    expires_in = fields.Char('Expires IN', readonly=True)
    subscription_type = fields.Char('Subscription Type', readonly=True)
    resource = fields.Char('Resource', readonly=True)
    user_id = fields.Many2one('res.users', default=lambda self: self.env.user)

class OfficeLeadModel(models.Model):
    _inherit = 'crm.lead'

    office_id = fields.Char('Office365 Id')