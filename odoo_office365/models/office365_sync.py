import logging
import re
from odoo.exceptions import ValidationError, UserError
from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools
import requests
from bs4 import BeautifulSoup
import base64
import threading
import json
from datetime import datetime
import time
from datetime import timedelta
import dateutil.parser as p
import traceback

_logger = logging.getLogger(__name__)
_image_dataurl = re.compile(r'(data:image/[a-z]+?);base64,([a-z0-9+/]{3,}=*)([\'"])', re.I)
_import_history = {}


class Office365(models.Model):
    """
    This class give functionality to user for Office365 Integration
    """
    _name = 'office.sync'
    _description = "Office365 - Connector"
    _rec_name = 'name'

    name = fields.Char("Name", compute='_get_name')
    event_del_flag = fields.Boolean('Delete events from Office365 calendar when delete in Odoo.',groups="base.group_user")
    office365_event_del_flag = fields.Boolean('Delete event from Odoo, if the event is deleted from Office 365.',groups="base.group_user")
    customers_count = fields.Integer("Customer Count", compute='compute_count')
    tasks_count = fields.Integer("Tasks Count", compute='compute_count')
    events_count = fields.Integer("Events Count", compute='compute_count')
    res_user = fields.Many2one('res.users', string='User', default=lambda self:self.env.user.id, readonly=True)
    is_manual_sync = fields.Boolean(string="Custom Date Range", )
    is_auto_sync = fields.Boolean(string="Auto Scheduler", )
    is_import_contact = fields.Boolean()
    is_import_email = fields.Boolean()
    is_import_calendar = fields.Boolean()
    is_import_task = fields.Boolean()
    configure_mail_server = fields.Boolean()
    is_export_contact = fields.Boolean()
    is_export_calendar = fields.Boolean()
    is_export_task = fields.Boolean()
    user_email = fields.Char("Email")
    user_password_secret = fields.Char("Password")
    from_date = fields.Datetime(string="From Date", required=False, )
    to_date = fields.Datetime(string="To Date", required=False, )
    is_manual_execute = fields.Boolean(string="Manual Execute", )
    categories = fields.Many2many('calendar.event.type', string='Select Event Category')
    contact_categories = fields.Many2many('res.partner.category', string='Select Contact Category')
    calendar_id = fields.Many2one(comodel_name="office.calendars", string="Office365 Calendars", required=False,)

    def compute_count(self):
        try:
            for record in self:
                record.customers_count = self.env['res.partner'].search_count([('office_contact_id', '!=', None),('create_uid', '=', self.env.user.id)])
                record.tasks_count = self.env['mail.activity'].search_count([('office_id', '!=', None),('create_uid', '=', self.env.user.id)])
                record.events_count = self.env['calendar.event'].search_count([('office_id', '!=', None),('user_id', '=', self.env.user.id)])
        except Exception as e:
            raise ValidationError(str(e))

    def get_code(self):
        try:
            if self.env.user.redirect_url and self.env.user.client_id and self.env.user.client_secret:
                return {
                    'name': 'login',
                    'view_id': False,
                    "type": "ir.actions.act_url",
                    'target': 'self',
                    'url': self.env.user.login_url
                }
            else:
                raise ValidationError('Office 365 Credentials are missing. Please! ask admin to add credentials')
        except Exception as e:
            raise ValidationError(str(e))

    def action_user_form_office_365(self):
        try:
            existing_configurations = self.env['office.sync'].sudo().search([
                ('res_user', '=', self.env.user.id)])
            context = dict(self._context)
            current_logged_in_uid = self._context.get('uid')
            # existing_configurations = self.search([('for_auto_user_id', '=', current_logged_in_uid)])
            data = {
                'name': 'Office365 Sync Mail and Contact',
                'res_model': 'office.sync',
                'target': 'current',
                'view_mode': 'form',
                'type': 'ir.actions.act_window'
            }
            if not existing_configurations:
                return data

            data['res_id'] = existing_configurations[0].id
            return data
        except Exception as e:
            raise ValidationError(str(e))

    def _get_name(self):
        try:
            for i in self:
                i.name = self.env.user.name + "'s " + "Configuration"
        except Exception as e:
            raise ValidationError(str(e))

    def previewScheduleAction(self):
        try:
            self.ensure_one()
            officeModel = self.env['ir.model'].search([('model', '=', 'office.sync')])

            return {
                'type': 'ir.actions.act_window',
                'name': 'Schedulers',
                'view_mode': 'tree',
                'res_model': 'ir.cron',
                'domain': [('model_id', '=', officeModel.id)],
                'context': "{'create': False}"
            }
        except Exception as e:
            raise ValidationError(str(e))

    def get_customers(self):
        try:
            self.ensure_one()
            return {
                'type': 'ir.actions.act_window',
                'name': 'Office 365 Customers',
                'view_mode': 'tree',
                'res_model': 'res.partner',
                'domain': [('office_contact_id', '!=', None),('create_uid', '=', self.env.user.id)],
                'context': "{'create': False}"
            }
        except Exception as e:
            raise ValidationError(str(e))

    def get_tasks(self):
        try:
            self.ensure_one()
            return {
                'type': 'ir.actions.act_window',
                'name': 'Office 365 Tasks',
                'view_mode': 'tree',
                'res_model': 'mail.activity',
                'domain': [('office_id', '!=', None),('create_uid', '=', self.env.user.id)],
                'context': "{'create': False}"
            }
        except Exception as e:
            raise ValidationError(str(e))

    def get_events(self):
        try:
            self.ensure_one()
            return {
                'type': 'ir.actions.act_window',
                'name': 'Office 365 Events',
                'view_mode': 'tree',
                'res_model': 'calendar.event',
                'domain': [('office_id', '!=', None),('user_id', '=', self.env.user.id)],
                'context': "{'create': False}"
            }
        except Exception as e:
            raise ValidationError(str(e))

    def insert_history_line(self, id, added, updated, is_manual, status, action):
        try:
            contacts_added = 0
            contacts_updated = 0
            if action == 'contacts':
                contacts_added = added
                contacts_updated = updated
            history = self.env['sync.history']
            history.create({'last_sync': datetime.now(),
                            'no_im_contact': contacts_added,
                            'no_up_contact': contacts_updated,
                            'sync_type': 'manual' if is_manual else 'auto',
                            'sync_id': id,
                            'no_up_task': 0,
                            'no_im_email': 0,
                            'no_im_calender': 0,
                            'no_up_calender': 0,
                            'no_im_task': 0,
                            'status': status if status else 'Success',
                            })
            self.env.cr.commit()
        except Exception as e:
            raise ValidationError(str(e))

    def execute_operation(self):
        try:
            self.checkTokenExpiryDate(self.res_user)
            self.env.user.event_del_flag = self.event_del_flag
            self.env.user.office365_event_del_flag = self.office365_event_del_flag
            if self.configure_mail_server:
                smtp_redundency = self.env['ir.mail_server'].search([('smtp_host', '=', 'smtp-mail.outlook.com')])
                if smtp_redundency:
                    if smtp_redundency.active == self.configure_mail_server:
                        pass
                    else:
                        smtp_redundency.write({
                            'active': self.configure_mail_server
                        })
                        self.env.cr.commit()
            if self.is_manual_sync:
                odooUser = self.res_user
                if odooUser.refresh_token:
                    if self.is_import_contact or self.is_import_task or self.is_import_email or self.is_import_calendar:
                        if self.is_manual_sync:
                            self.import_data(False)

                    if self.is_export_task or self.is_export_contact or self.is_export_calendar:
                        if self.is_manual_sync:
                            self.export_data(False)

                    if not self.is_import_contact and not self.is_import_task and not self.is_import_email \
                            and not self.is_import_calendar and not self.is_export_task and not self.is_export_contact \
                            and not self.is_export_calendar and not self.event_del_flag and not self.office365_event_del_flag:
                        raise ValidationError('No operation utility is selected.')
                else:
                    raise ValidationError('Please login with your Office 365 account to perform operation')
            else:
                raise ValidationError('Please specify the type of operation.')
        except Exception as e:
            raise ValidationError(str(e))

    def auto_import(self):
        try:
            activeAutoImportSettings = self.env['office.sync'].search([('is_auto_sync', '=', True)])
            for activeAutoImportSetting in activeAutoImportSettings:
                self.import_data(True, activeAutoImportSetting)
        except Exception as e:
            raise ValidationError(str(e))

    def auto_export(self):
        try:
            activeAutoExportSettings = self.env['office.sync'].search([('is_auto_sync', '=', True)])
            for activeAutoExportSetting in activeAutoExportSettings:
                self.export_data(True, activeAutoExportSetting)
        except Exception as e:
            raise ValidationError(str(e))

    def auto_configure_webhook(self):
        activeAutoExportSettings = self.env['office.sync'].search([('id', '!=', None)])
        for activeAutoExportSetting in activeAutoExportSettings:
            self.configure_webhook(True, activeAutoExportSetting)

    def import_data(self, auto, setting=None):
        try:
            data_dictionary = {}
            status = 'import'
            sync_type = 'manual'
            if auto:
                self = setting
                sync_type = 'scheduled'

            self.checkTokenExpiryDate(self.res_user)

            if self.is_import_contact:
                data_dictionary["importContacts"] = self.import_contacts()
            if self.is_import_task:
                print(self.is_import_task)
                data_dictionary["importTasks"] = self.import_tasks()
            if self.is_import_email:
                print(self.is_import_email)
                data_dictionary["importEmails"] = self.sync_customer_mail()
            if self.is_import_calendar:
                data_dictionary["importCalendars"] = self.import_calendar(auto)

            self.env['sync.history'].create({
                'user_id' : self.res_user.id,
                'last_sync' : datetime.now(),
                'numberContacts' : data_dictionary['importContacts']['importedContacts'] if 'importContacts' in data_dictionary else None,
                'numberContactsUpdated' : data_dictionary['importContacts']['updatedContacts'] if 'importContacts' in data_dictionary else None,
                'numberEmails' : data_dictionary['importEmails']['importedEmails'] if 'importEmails' in  data_dictionary else None,
                'numberTasks' : data_dictionary['importTasks']['importedTasks'] if 'importTasks' in data_dictionary else None,
                'numberTasksUpdated' : data_dictionary['importTasks']['updatedTasks'] if 'importTasks' in data_dictionary else None,
                'numberCalendars' : data_dictionary['importCalendars']['importedCalendars'] if 'importCalendars' in data_dictionary else None,
                'numberCalendarsUpdated' : data_dictionary['importCalendars']['updatedCalendars'] if 'importCalendars' in data_dictionary else None,
                'status': status,
                'sync_type': sync_type,
            })
            self.env.cr.commit()
        except Exception as e:
            raise ValidationError(str(e))

    def export_data(self, auto, setting=None):
        try:
            data_dictionary = {}

            status = 'export'
            sync_type = 'manual'

            if auto:
                self = setting
                sync_type = 'scheduled'

            self.checkTokenExpiryDate(self.res_user)

            if self.is_export_task:
                data_dictionary["exportTasks"] = self.export_tasks()
            if self.is_export_contact:
                data_dictionary["exportContacts"] = self.export_contacts()
            if self.is_export_calendar:
                data_dictionary["exportCalendars"] = self.export_calendar()

            self.env['sync.history'].create({
                'user_id' : self.res_user.id,
                'last_sync' : datetime.now(),
                'numberContacts' : data_dictionary['exportContacts']['exportedContacts'] if 'exportContacts' in data_dictionary else None,
                'numberContactsUpdated' : data_dictionary['exportContacts']['updatedContacts'] if 'exportContacts' in data_dictionary  else None,
                'numberEmails' : data_dictionary['exportEmails']['exportedEmails'] if 'exportEmails' in  data_dictionary else None,
                'numberTasks' : data_dictionary['exportTasks']['exportedTasks'] if 'exportTasks' in data_dictionary else None,
                'numberTasksUpdated' : data_dictionary['exportTasks']['updatedTasks'] if 'exportTasks' in data_dictionary else None,
                'numberCalendars' : data_dictionary['exportCalendars']['exportedCalenders'] if 'exportCalendars' in data_dictionary else None,
                'numberCalendarsUpdated' : data_dictionary['exportCalendars']['updatedCalenders'] if 'exportCalendars' in data_dictionary else None,
                'status' : status,
                'sync_type': sync_type,
                })
            self.env.cr.commit()
        except Exception as e:
            raise ValidationError(str(e))

    ''' 
            These following methods are responsible for managing webhook Subscription
        '''

    def configure_webhook(self, auto=None, setting=None):
        try:
            if auto == True:
                self = setting
            if self.env.user.token:
                self.checkTokenExpiryDate(self.env.user)
                webhook = self.env['office.webhook'].sudo().search([('name', '=', 'mail')])
                if not webhook:
                    if self.env.user.webhook_redirect_url:
                        headers = {
                            'Authorization': 'Bearer {0}'.format(self.env.user.token),
                            'Content-type': 'application/json',
                        }
                        payload = {
                            "changeType": "created",
                            "notificationUrl": self.env.user.webhook_redirect_url + "/mail_webhook",
                            "resource": "/me/mailfolders('inbox')/messages",
                            "expirationDateTime": str(datetime.now().isoformat().split('T')[0]) + 'T23:59:00.0000000Z',
                        }
                        subscription_url = "https://graph.microsoft.com/v1.0/subscriptions"
                        subscription_response = requests.post(url=subscription_url, data=json.dumps(payload),
                                                              headers=headers)
                        if subscription_response.status_code == 201:
                            subcscription = json.loads(subscription_response.content.decode('utf-8'))
                            self.env['office.webhook'].create({
                                "name": 'mail',
                                "subscription_id": subcscription['id'],
                                "office_webhook_change_type": subcscription['changeType'],
                                "expires_in": subcscription['expirationDateTime'],
                                "subscription_type": subcscription['changeType'],
                                "resource": subcscription['resource'],
                                "user_id": self.res_user.id
                            })
                            self.env.cr.commit()
                            print("done created the webhook")
                    else:
                        raise ValidationError('Kindly provide Webhook Redirect Url')

                else:
                    self.checkSubscriptionExpiryDate(webhook)

        except Exception as e:
            raise ValidationError(str(e))

    def checkSubscriptionExpiryDate(self, webhook_subscription):
        try:
            if webhook_subscription.expires_in:
                expires_in = datetime.strptime(
                    datetime.strftime(p.parse(webhook_subscription.expires_in), "%Y-%m-%dT%H:%M:%S"),
                    "%Y-%m-%dT%H:%M:%S")
                now = datetime.now() + timedelta(minutes=60)
                if now > expires_in:
                    self.update_subscription(webhook_subscription)
        except Exception as e:
            raise ValidationError(str(e))

    def update_subscription(self, webhook_subscription):
        expires_in = datetime.strptime(datetime.strftime(p.parse(webhook_subscription.expires_in), "%Y-%m-%dT%H:%M:%S"),
                                       "%Y-%m-%dT%H:%M:%S")
        expires_inn = expires_in + timedelta(days=1)
        headers = {
            'Authorization': 'Bearer {0}'.format(self.res_user.token),
            'Content-type': 'application/json',
        }
        payload = {
            # "includeResourceData": True,
            "expirationDateTime": expires_inn.isoformat() + 'Z',
            # "notificationContentType": "text/plain"
        }
        subscription_url = "https://graph.microsoft.com/v1.0/subscriptions/" + webhook_subscription.subscription_id + ""
        subscription_update_response = requests.patch(url=subscription_url, data=json.dumps(payload), headers=headers)
        if subscription_update_response.status_code == 200:
            subcscription = json.loads(subscription_update_response.content.decode('utf-8'))
            webhook_subscription.sudo().write({
                "expires_in": subcscription['expirationDateTime'],
            })
            self.env.cr.commit()
        else:
            raise _logger.error(
                str("Your Token has been expired and while updating automatically we are facing issue can you please "
                    "check your credentials or again login with your Office 365 account"))

    ''' 
    These following methods are responsible fro importing contacts from Office 365
    '''

    def import_contacts(self):
        try:
            office_contacts = 0
            update_contact = 0
            if self.contact_categories:
                if not self.from_date and not self.to_date:
                    for catg in self.contact_categories:
                        url = 'https://graph.microsoft.com/v1.0/me/contacts?$filter=categories/any(a:a+eq+\'{}\')'.format(
                            catg.name.replace(' ', '+'))
                        office_contacts, update_contact = self.create_contacts(url)
                elif self.from_date and self.to_date:
                    for catg in self.contact_categories:
                        url = 'https://graph.microsoft.com/v1.0/me/' \
                              'contacts?$filter=lastModifiedDateTime ge {} and lastModifiedDateTime le {}'.format(
                            self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                            self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ") + " and categories/any(a:a+eq+'{}')".format(
                                catg.name.replace(' ', '+')))
                        office_contacts, update_contact = self.create_contacts(url)
                else:
                    raise ValidationError('Please select proper date range i.e. you have to select from and to date or '
                                          'else leave them blank')
            else:
                if not self.from_date and not self.to_date:
                    url = 'https://graph.microsoft.com/v1.0/me/contacts'
                    office_contacts, update_contact = self.create_contacts(url)
                elif self.from_date and self.to_date:
                    url = 'https://graph.microsoft.com/v1.0/me/' \
                          'contacts?$filter=lastModifiedDateTime ge {} and lastModifiedDateTime le {}'.format(
                        self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                        self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                    office_contacts, update_contact = self.create_contacts(url)
                else:
                    raise ValidationError('Please select proper date range i.e. you have to select from and to date or '
                                          'else leave them blank')
            import_dictionary = {
                'importedContacts': office_contacts,
                'updatedContacts': update_contact
            }

            return import_dictionary
        except Exception as e:
            raise ValidationError(str(e))

    def create_contacts(self, url):
        try:
            office_contacts = []
            update_contact = []
            headers = {
                'Authorization': 'Bearer {0}'.format(self.res_user.token),
                'Accept': 'application/json',
                'connection': 'keep-Alive'}

            while True:
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    response = json.loads(response.content.decode('utf-8'))
                    if 'value' in response:
                        for each_contact in response['value']:
                            response_image_data = ""
                            odoo_customer = self.env['res.partner'].search([('office_contact_id', '=', each_contact['id'])])
                            name = None
                            categories_ids = None
                            if each_contact['displayName']:
                                firstname = each_contact['givenName'] if each_contact['givenName'] else ''
                                middlename = each_contact['middleName'] if each_contact['middleName'] else ''
                                lastname = each_contact['surname'] if each_contact['surname'] else ''
                                name = firstname + " " + middlename + " " + lastname
                                company_type = 'person'
                                if len(each_contact['homePhones']) > 0:
                                    phone = each_contact['homePhones'][0]
                            else:
                                name = each_contact['companyName']
                                company_type = 'company'
                                if len(each_contact['businessPhones']) > 0:
                                    phone = each_contact['businessPhones'][0]
                            if len(each_contact['categories']) > 0:
                                categories_ids = self.getContactsOdooCategory(each_contact['categories'])

                            officeModifiedDate = datetime.strptime(
                                datetime.strftime(p.parse(each_contact['lastModifiedDateTime']), "%Y-%m-%dT%H:%M:%S"),
                                "%Y-%m-%dT%H:%M:%S")

                            if not odoo_customer:
                                url_image = "https://graph.microsoft.com/v1.0/me/contacts/"+each_contact['id']+"/photo/$value"
                                image_headers = {
                                    'Authorization': 'Bearer {0}'.format(self.res_user.token),
                                    'Accept': "image/jpeg",
                                    'connection': 'keep-Alive'}
                                response_image = requests.get(url_image, headers=image_headers)
                                if response_image.status_code == 200:
                                    response_image_data = base64.b64encode(response_image.content)
                                odoo_cust = self.env['res.partner'].create({
                                    'office_contact_id': each_contact['id'],
                                    'modified_date': officeModifiedDate,
                                    'company_id': self.res_user.company_id.id,
                                    'company_type': company_type,
                                    'name': name if name else ' ',
                                    'image_1920': response_image_data,
                                    'firstName': each_contact['givenName'] if each_contact['givenName'] else '',
                                    'lastName': each_contact['surname'] if each_contact['surname'] else '',
                                    'middleName': each_contact['middleName'] if each_contact['middleName'] else '',
                                    'function': each_contact['jobTitle'] if each_contact['jobTitle'] else None,
                                    'phone': each_contact['businessPhones'][0] if each_contact['businessPhones'] else None,
                                    'mobile': each_contact['mobilePhone'] if each_contact['mobilePhone'] else None,
                                    'email': each_contact['emailAddresses'][0]['address'].lower() if each_contact['emailAddresses'] else None,
                                    'website': each_contact['businessHomePage'] if each_contact['businessHomePage'] else None,
                                    'location': 'Office365 Contact',
                                    'category_id': [[6, 0, categories_ids]] if categories_ids else None,
                                })
                                self.env.cr.commit()
                                self.env.cr.execute( """UPDATE res_partner SET create_uid='%s' WHERE id=%s""" % (self.res_user.id,odoo_cust.id))
                                self.env.cr.commit()
                                if each_contact['homeAddress']:
                                    country = self.env['res.country'].search([('name', '=', each_contact['homeAddress']['countryOrRegion'])]) if 'country' in each_contact['homeAddress'] else None
                                    state = self.env['res.country.state'].search([('name', '=', each_contact['homeAddress']['state']), ('country_id', '=', country)]) if 'state' in each_contact['homeAddress'] else None
                                    delivery_address=self.env['res.partner'].create({
                                        'street': each_contact['homeAddress']['street'] if 'street' in each_contact['homeAddress'] else None,
                                        'city': each_contact['homeAddress']['city'] if 'city' in each_contact['homeAddress'] else None,
                                        'state_id': state.id if state else None,
                                        'country_id': country.id if country else None,
                                        'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in each_contact['homeAddress'] else None,
                                        'parent_id': odoo_cust.id,
                                        'type': 'delivery',
                                    })
                                    self.env.cr.commit()
                                    self.env.cr.execute("""UPDATE res_partner SET create_uid='%s' WHERE id=%s""" % (
                                    self.res_user.id, delivery_address.id))
                                    self.env.cr.commit()
                                if each_contact['businessAddress']:
                                    country = self.env['res.country'].search([('name', '=', each_contact['businessAddress']['countryOrRegion'])]) if 'country' in each_contact['businessAddress'] else None
                                    state = self.env['res.country.state'].search([('name', '=', each_contact['businessAddress']['state']), ('country_id', '=', country)]) if 'state' in each_contact['businessAddress'] else None
                                    invoice_address=self.env['res.partner'].create({
                                        'street': each_contact['businessAddress']['street'] if 'street' in each_contact['businessAddress'] else None,
                                        'city': each_contact['businessAddress']['city'] if 'city' in each_contact['businessAddress'] else None,
                                        'state_id': state.id if state else None,
                                        'country_id': country.id if country else None,
                                        'zip': each_contact['businessAddress']['postalCode'] if 'postalCode' in each_contact['businessAddress'] else None,
                                        'parent_id': odoo_cust.id,
                                        'type': 'invoice',
                                    })
                                    self.env.cr.commit()
                                    self.env.cr.execute("""UPDATE res_partner SET create_uid='%s' WHERE id=%s""" % (self.res_user.id, invoice_address.id))
                                    self.env.cr.commit()
                                if each_contact['otherAddress']:
                                    country = self.env['res.country'].search([('name', '=', each_contact['otherAddress']['countryOrRegion'])]) if 'country' in each_contact['businessAddress'] else None
                                    state = self.env['res.country.state'].search([('name', '=', each_contact['otherAddress']['state']), ('country_id', '=', country)]) if 'state' in each_contact['otherAddress'] else None
                                    other_address=self.env['res.partner'].create({
                                        'street': each_contact['otherAddress']['street'] if 'street' in each_contact['otherAddress'] else None,
                                        'city': each_contact['otherAddress']['city'] if 'city' in each_contact['otherAddress'] else None,
                                        'state_id': state.id if state else None,
                                        'country_id': country.id if country else None,
                                        'zip': each_contact['otherAddress']['postalCode'] if 'postalCode' in each_contact['otherAddress'] else None,
                                        'parent_id': odoo_cust.id,
                                        'type': 'other',
                                    })
                                    self.env.cr.commit()
                                    self.env.cr.execute("""UPDATE res_partner SET create_uid='%s' WHERE id=%s""" % (
                                    self.res_user.id, other_address.id))
                                    self.env.cr.commit()
                                office_contacts.append(odoo_cust.id)
                            else:
                                if odoo_customer.modified_date:
                                    if odoo_customer.modified_date >= officeModifiedDate:
                                        continue
                                    else:
                                        odoo_customer.write({
                                            'office_contact_id': each_contact['id'],
                                            'modified_date': officeModifiedDate,
                                            'company_id': self.res_user.company_id.id,
                                            'company_type': company_type,
                                            'name': name,
                                            'firstName': each_contact['givenName'] if each_contact['givenName'] else '',
                                            'lastName': each_contact['surname'] if each_contact['surname'] else '',
                                            'middleName': each_contact['middleName'] if each_contact['middleName'] else '',
                                            'function': each_contact['jobTitle'] if each_contact['jobTitle'] else None,
                                            'phone': each_contact['businessPhones'][0] if each_contact['businessPhones'] else None,
                                            'mobile': each_contact['mobilePhone'] if each_contact['mobilePhone'] else None,
                                            'email': each_contact['emailAddresses'][0]['address'].lower() if each_contact['emailAddresses'] else None,
                                            'website': each_contact['businessHomePage'] if each_contact['businessHomePage'] else None,
                                            'location': 'Office365 Contact',
                                            'category_id': [[6, 0, categories_ids]] if categories_ids else None,
                                        })
                                        for child in odoo_customer.child_ids:
                                            if child.type == 'delivery':
                                                country = self.env['res.country'].search([('name', '=',each_contact['homeAddress']['countryOrRegion'])]) if 'country' in each_contact['homeAddress'] else None
                                                state = self.env['res.country.state'].search([('name', '=', each_contact['homeAddress']['state']),('country_id', '=', country)]) if 'state' in each_contact['homeAddress'] else None
                                                child.write({
                                                    'street': each_contact['homeAddress']['street'] if 'street' in each_contact[ 'homeAddress'] else None,
                                                    'city': each_contact['homeAddress']['city'] if 'city' in each_contact[ 'homeAddress'] else None,
                                                    'state_id': state.id if state else None,
                                                    'country_id': country.id if country else None,
                                                    'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in each_contact['homeAddress'] else None,
                                                    'parent_id': odoo_customer.id,
                                                    'type': 'delivery',
                                                })
                                            elif child.type == 'invoice':
                                                country = self.env['res.country'].search([('name', '=',each_contact['businessAddress']['countryOrRegion'])]) if 'country' in each_contact['businessAddress'] else None
                                                state = self.env['res.country.state'].search([('name', '=', each_contact['businessAddress']['state']),('country_id', '=', country)]) if 'state' in each_contact['businessAddress'] else None
                                                child.write({
                                                    'street': each_contact['businessAddress']['street'] if 'street' in each_contact[ 'businessAddress'] else None,
                                                    'city': each_contact['businessAddress']['city'] if 'city' in each_contact[ 'businessAddress'] else None,
                                                    'state_id': state.id if state else None,
                                                    'country_id': country.id if country else None,
                                                    'zip': each_contact['businessAddress']['postalCode'] if 'postalCode' in each_contact['businessAddress'] else None,
                                                    'parent_id': odoo_customer.id,
                                                    'type': 'invoice',
                                                })
                                            elif child.type == 'other':
                                                country = self.env['res.country'].search([('name', '=',each_contact['otherAddress']['countryOrRegion'])]) if 'country' in each_contact['otherAddress'] else None
                                                state = self.env['res.country.state'].search([('name', '=', each_contact['otherAddress']['state']),('country_id', '=', country)]) if 'state' in each_contact['otherAddress'] else None
                                                child.write({
                                                    'street': each_contact['otherAddress']['street'] if 'street' in each_contact[ 'otherAddress'] else None,
                                                    'city': each_contact['otherAddress']['city'] if 'city' in each_contact[ 'otherAddress'] else None,
                                                    'state_id': state.id if state else None,
                                                    'country_id': country.id if country else None,
                                                    'zip': each_contact['otherAddress']['postalCode'] if 'postalCode' in each_contact['otherAddress'] else None,
                                                    'parent_id': odoo_customer.id,
                                                    'type': 'other',
                                                })
                                        update_contact.append(odoo_customer.id)
                            self.env.cr.commit()
                        if '@odata.nextLink' in response:
                            url = response['@odata.nextLink']
                        else:
                            break
                else:
                    raise ValidationError(("Unable to connect with office.Kindly build connection again"))
                    break
            return len(office_contacts), len(update_contact)
        except Exception as e:
            raise ValidationError(str(e))

    def export_contacts(self):

        new_contact = []
        update_contact = []
        skip_ids = []
        try:
            res_user = self.res_user
            if self.from_date and not self.to_date:
                raise ValidationError('Warning!', 'Please! Select "To Date" to Import Events.')
            if not self.from_date and self.to_date:
                raise ValidationError('Warning!', 'Please! Select "From Date" to Import Events.')
            if self.from_date > self.to_date:
                raise ValidationError('Warning!', 'Please! Enter Date in correct Order !')
            if self.from_date and self.to_date:
                from_date = self.from_date
                to_date = self.to_date

            odoo_contacts = self.env['res.partner'].search([("create_uid", '=', res_user.id)])
            users = self.env['res.users'].sudo().search([('id', '!=', None)])
            companies = self.env['res.company'].sudo().search([('id', '!=', None)])
            for company in companies:
                skip_ids.append(company.partner_id.id)
            for user in users:
                skip_ids.append(user.partner_id.id)
            if self.from_date and self.to_date:
                odoo_contacts = odoo_contacts.search([('write_date', '>=', self.from_date), ('write_date', '<=', self.to_date)])

            headers = {
                'Host': 'outlook.office365.com',
                'Authorization': 'Bearer {0}'.format(res_user.token),
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'connection': 'keep-Alive'
            }

            for contact in odoo_contacts:
                if contact.type == 'delivery' or contact.type == 'invoice' or contact.type == 'other' or contact.id in skip_ids:
                    pass
                else:
                    categories = []
                    if contact.category_id:
                        for category in contact.category_id:
                            categories.append(category.name)
                    data = {
                        'displayName': contact.name if contact.is_company == False else None,
                        'givenName': contact.firstName if contact.firstName else None,
                        'surname': contact.lastName if contact.lastName else None,
                        'categories': categories,
                        'companyName': contact.company_name if contact.is_company == False and contact.company_name == True else contact.name if contact.is_company == True else None,
                        'jobTitle': contact.function if contact.function else None,
                        'businessPhones': [contact.phone if contact.phone else None,
                        ],
                        'mobilePhone': contact.mobile if contact.mobile else None,
                        'businessHomePage': contact.website if contact.website else None,
                    }
                    if contact.email:
                        data["emailAddresses"] = [
                            {
                                "address": contact.email if contact.email else None,
                            },
                        ]
                    for child in contact.child_ids:
                        if child.type == 'delivery':
                            data["homeAddress"] = {
                                "street": child.street if child.street else (
                                    child.street2 if child.street2 else None),
                                "city": child.city if child.city else None,
                                "state": child.state_id.name if child.state_id else None,
                                "countryOrRegion": child.country_id.name if child.country_id else None,
                                "postalCode": child.zip if child.zip else None,
                            }
                        elif child.type == 'invoice':
                            data["businessAddress"] = {
                                "street": child.street if child.street else (
                                    child.street2 if child.street2 else None),
                                "city": child.city if child.city else None,
                                "state": child.state_id.name if child.state_id else None,
                                "countryOrRegion": child.country_id.name if child.country_id else None,
                                "postalCode": child.zip if child.zip else None,
                            }
                        elif child.type == 'other':
                            data["otherAddress"] = {
                                "street": child.street if child.street else (
                                    child.street2 if child.street2 else None),
                                "city": child.city if child.city else None,
                                "state": child.state_id.name if child.state_id else None,
                                "countryOrRegion": child.country_id.name if child.country_id else None,
                                "postalCode": child.zip if child.zip else None,
                            }

                    if contact.office_contact_id:
                        url = "https://graph.microsoft.com/v1.0/me/contacts/{}".format(
                            contact.office_contact_id)
                        response = requests.get(url,headers={
                                'Host': 'outlook.office.com',
                                'Authorization': 'Bearer {0}'.format(res_user.token),
                                'Accept': 'application/json',
                                'X-Target-URL': 'http://outlook.office.com',
                                'connection': 'keep-Alive'
                            })
                        if response.status_code == 200 or response.status_code == 201:
                            if contact.create_date < contact.write_date:
                                res_event = json.loads((response.content.decode('utf-8')))
                                modified_at = datetime.strptime(res_event['lastModifiedDateTime'][:-1], '%Y-%m-%dT%H:%M:%S')
                                if modified_at > contact.write_date or modified_at == contact.write_date:
                                    pass
                                else:
                                    update_response = requests.patch(
                                        'https://graph.microsoft.com/v1.0/me/contacts/' + str(
                                            contact.office_contact_id), data=json.dumps(data), headers=headers)
                                    if update_response.status_code == 200 or update_response.status_code == 201:
                                        update_response = json.loads(update_response.content.decode('utf-8'))
                                        contact.write({'office_contact_id': update_response['id']})
                                        update_contact.append(update_response['id'])
                                    # contact.is_update = False
                                    else:
                                        update_response = json.loads(update_response.content.decode('utf-8'))
                                        self.env['office.export.exceptions'].create({
                                            'last_sync': datetime.now(),
                                            'endpoint': 'Contact',
                                            'operation': 'update',
                                            'contact_id':contact.id,
                                            'event_id': None,
                                            'task_id': None,
                                            'skip': True,
                                            'description': update_response['error']['code']+' '+update_response['error']['message'],
                                            'model': 'res.partner',
                                            # 'sync_type': 'scheduled' if Auto else 'manual',
                                            'user_id': self.res_user.id,
                                        })
                                        self.env.cr.commit()
                                        continue
                        else:
                            post_response = requests.post('https://graph.microsoft.com/v1.0/me/contacts', data=json.dumps(data),headers=headers)
                            if post_response.status_code == 200 or post_response.status_code == 201:
                                post_response = json.loads(post_response.content.decode('utf-8'))
                                contact.write({'office_contact_id': post_response['id']})
                                contact.is_update = False
                                new_contact.append(post_response['id'])
                            else:
                                update_response = json.loads(post_response.content.decode('utf-8'))
                                self.env['office.export.exceptions'].create({
                                    'last_sync': datetime.now(),
                                    'endpoint': 'Contact',
                                    'operation': 'update',
                                    'contact_id': contact.id,
                                    'event_id': None,
                                    'task_id': None,
                                    'skip': True,
                                    'description': update_response['error']['code'] + ' ' + update_response['error'][
                                        'message'],
                                    'model': 'res.partner',
                                    # 'sync_type': 'scheduled' if Auto else 'manual',
                                    'user_id': self.res_user.id,
                                })
                                self.env.cr.commit()
                                continue
                    else:
                        post_response = requests.post(
                            'https://graph.microsoft.com/v1.0/me/contacts', data=json.dumps(data),
                            headers=headers
                        )

                        if post_response.status_code == 200 or post_response.status_code == 201:
                            post_response = json.loads(post_response.content.decode('utf-8'))
                            contact.write({'office_contact_id': post_response['id']})
                            contact.is_update = False
                            new_contact.append(post_response['id'])
                        else:
                            update_response = json.loads(post_response.content.decode('utf-8'))
                            self.env['office.export.exceptions'].create({
                                'last_sync': datetime.now(),
                                'endpoint': 'Contact',
                                'operation': 'update',
                                'contact_id': contact.id,
                                'event_id': None,
                                'task_id': None,
                                'skip': True,
                                'description': update_response['error']['code'] + ' ' + update_response['error'][
                                    'message'],
                                'model': 'Event',
                                # 'sync_type': 'scheduled' if Auto else 'manual',
                                'user_id': self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue

            export_dictionary = {
                'exportedContacts': len(new_contact),
                'updatedContacts': len(update_contact)
            }

            return export_dictionary

        except Exception as e:
            raise ValidationError(_(str(e)))

    ''' 
        These methods are responsible fro importing calendars from Office 365
    '''

    def import_calendar(self,Auto):
        try:
            office_calendars = 0
            update_calendars = 0
            if self.categories:
                if not self.from_date and not self.to_date:
                    for catg in self.categories:
                        if self.calendar_id and self.calendar_id.calendar_id:
                            url = "https://graph.microsoft.com/v1.0/me/calendars/" + str(
                                self.calendar_id.calendar_id) + "/events?$filter=categories/any(a:a+eq+'{}')".format(
                                catg.name.replace(' ', '+'))
                        else:
                            url = "https://graph.microsoft.com/v1.0/me/events?$filter=categories/any(a:a+eq+'{}')".format(
                                catg.name.replace(' ', '+'))
                        office_calendars, update_calendars = self.create_events(url,Auto)
                if self.from_date and self.to_date:
                    categ_name = []
                    for catg in self.categories:
                        if self.calendar_id:
                            url = "https://graph.microsoft.com/v1.0/me/calendars/" + str(
                                self.calendar_id.calendar_id) + "/events?$filter=lastModifiedDateTime%20gt%20{}%20and%20lastModifiedDateTime%20lt%20{}".format(
                                self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"), self.to_date.strftime(
                                    "%Y-%m-%dT%H:%M:%SZ")) + " and categories/any(a:a+eq+'{}')".format(
                                catg.name.replace(' ', '+'))
                        else:
                            url = "https://graph.microsoft.com/v1.0/me/calendars/events?$filter=lastModifiedDateTime%20gt%20{}%20and%20lastModifiedDateTime%20lt%20{}".format(
                                self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"), self.to_date.strftime(
                                    "%Y-%m-%dT%H:%M:%SZ")) + " and categories/any(a:a+eq+'{}')".format(
                                catg.name.replace(' ', '+'))

                        office_calendars, update_calendars = self.create_events(url,Auto)
            else:
                if self.from_date and self.to_date:
                    if self.calendar_id:
                        url = 'https://graph.microsoft.com/v1.0/me/calendars/' + str(
                            self.calendar_id.calendar_id) + '/events?$filter=lastModifiedDateTime%20gt%20{}%20and%20lastModifiedDateTime%20lt%20{}' \
                                  .format(self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                                          self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                    else:
                        url = 'https://graph.microsoft.com/v1.0/me/events?$filter=lastModifiedDateTime%20gt%20{}%20and%20lastModifiedDateTime%20lt%20{}' \
                            .format(self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                                    self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                    office_calendars, update_calendars = self.create_events(url,Auto)

                if not self.from_date and not self.to_date:
                    if self.calendar_id and self.calendar_id.calendar_id:
                        url = 'https://graph.microsoft.com/v1.0/me/calendars/' + str(
                            self.calendar_id.calendar_id) + '/events'
                    else:
                        url = 'https://graph.microsoft.com/v1.0/me/events'

                    office_calendars, update_calendars = self.create_events(url,Auto)


            import_dictionary = {
                'importedCalendars': office_calendars,
                'updatedCalendars': update_calendars,
            }

            return import_dictionary

        except Exception as e:
            traceback.print_exc()
            raise ValidationError(str(e))

    def create_events(self, url,Auto):
        odoo_event_ids = []
        update_event = []
        new_event = []
        try:
            header = {
            'Authorization': 'Bearer {0}'.format(self.res_user.token),
            'Accept': 'application/json',
            'connection': 'keep-Alive'}
            while True:
                response = requests.get(url, headers=header)
                if response.status_code == 200:
                    responseDecoded = json.loads((response.content.decode('utf-8')))
                    events = responseDecoded['value']
                    for event in events:
                        if self.res_user.office365_event_del_flag and self.from_date and self.to_date:
                            if event['id'] in odoo_event_ids:
                                odoo_event_ids.remove(event['id'])
                    for event in events:
                        try:
                            print('Calender: ' + event['subject'])
                            week_days = []
                            week_index = []
                            end_date = None
                            occurences = None
                            end_type = None
                            month_by = None
                            day = None
                            start_date=None
                            interval = None
                            odooEvent = self.env['calendar.event'].search([("office_id", "=", event['id'])])
                            categories_ids = None
                            if event['categories']:
                                categories_ids = self.getEventOdooCategory(event)
                            rRule = None
                            partner_ids = []
                            if event['attendees']:
                                partner_ids = self.getEventOdooAttendees(event['attendees'])
                                partner_ids.append(self.res_user.partner_id.id)
                            else:
                                partner_ids.append(self.res_user.partner_id.id)

                            location = event['location']['displayName'] if 'displayName' in event['location'] else None
                            startDateTime = datetime.strptime(datetime.strftime(p.parse(event['start']['dateTime']), "%Y-%m-%dT%H:%M:%S"),"%Y-%m-%dT%H:%M:%S")
                            endDateTime = datetime.strptime(datetime.strftime(p.parse(event['end']['dateTime']), "%Y-%m-%dT%H:%M:%S"),"%Y-%m-%dT%H:%M:%S")
                            officeModifiedDate = datetime.strptime(datetime.strftime(p.parse(event['lastModifiedDateTime']), "%Y-%m-%dT%H:%M:%S"),"%Y-%m-%dT%H:%M:%S")
                            if 'sensitivity' in event:
                                if event['sensitivity'] == 'private':
                                    privacy = 'private'
                                else:
                                    privacy = 'public'
                            if event['recurrence']:
                                interval = event['recurrence']['pattern']['interval'] if 'pattern' in event['recurrence'] else None
                                start_date = event['recurrence']['range']['startDate'] if 'range' in event['recurrence'] else None
                                if 'numbered' in event['recurrence']['range']['type']:
                                    end_type = 'count'
                                    occurences = event['recurrence']['range']['numberOfOccurrences'] if 'numberOfOccurrences' in event['recurrence']['range'] else None
                                if 'endDate' in event['recurrence']['range']['type']:
                                    end_type = 'end_date'
                                    end_date = event['recurrence']['range']['endDate']
                                if 'noEnd' in event['recurrence']['range']['type']:
                                    end_type = 'forever'
                                    #end_date = event['recurrence']['range']['endDate']
                                if 'daily' in event['recurrence']['pattern']['type']:
                                    rRule = 'daily'
                                elif 'weekly' in event['recurrence']['pattern']['type']:
                                    rRule = 'weekly'
                                elif 'absoluteMonthly' in event['recurrence']['pattern']['type'] or 'relativeMonthly' in event['recurrence']['pattern']['type']:
                                    rRule = 'monthly'
                                    if 'absoluteMonthly' in event['recurrence']['pattern']['type']:
                                        month_by = 'date'
                                        day = event['recurrence']['pattern']['dayOfMonth'] if 'dayOfMonth' in event['recurrence']['pattern'] else None
                                    else:
                                        month_by = 'day'
                                        indexes_dict={'first': '1', 'second': '2', 'third': '3', 'fourth': '4','last': '-1'}
                                        for index in indexes_dict:
                                            if event['recurrence']['pattern']['index'] == index:
                                                week_index.append(indexes_dict[index])
                                        week_dict = {'monday': 'MON', 'tuesday': 'TUE', 'wednesday': 'WED', 'thursday': 'THU','friday': 'FRI','saturday': 'SAT', 'sunday': 'SUN'}
                                        for week_day in week_dict:
                                            for outlook_day in event['recurrence']['pattern']['daysOfWeek']:
                                                if outlook_day == week_day:
                                                    week_days.append(week_dict[week_day])

                                elif 'absoluteYearly' in event['recurrence']['pattern']['type'] or 'relativeYearly' in event['recurrence']['pattern']['type']:
                                    rRule = 'yearly'
                            if not odooEvent:
                                try:
                                    odooEvent = self.env['calendar.event'].sudo().create({
                                        'office_id': event['id'],
                                        'name': event['subject'],
                                        'calendar_id': self.calendar_id.id if self.calendar_id else None,
                                        'category_name': event['categories'][0] if event['categories'] else None,
                                        "description": event['body']['content'],
                                        'privacy': privacy,
                                        'location': location,
                                        'start': startDateTime,
                                        'stop': endDateTime,
                                        'start_date': start_date if start_date else None,
                                        'until': end_date if end_date else None,
                                        'allday': event['isAllDay'],
                                        'categ_ids': [[6, 0, categories_ids]] if categories_ids else None,
                                        'show_as': event['showAs'],
                                        'recurrency': True if event['recurrence'] else False,
                                        'interval': interval,
                                        'end_type': end_type,
                                        'rrule_type': rRule,
                                        'count': occurences,
                                        'mon': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'monday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'tue': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'tuesday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'wed': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'wednesday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'thu': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'thursday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'fri': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'friday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'sat': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'saturday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'sun': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                            'pattern'].keys() and 'sunday' in event['recurrence']['pattern'][
                                                          'daysOfWeek'] else False,
                                        'month_by': month_by,
                                        'day': day,
                                        'byday': week_index[0] if week_index else None ,
                                        'weekday': week_days[0] if week_days else None ,
                                        'partner_ids': [[6, 0, partner_ids]] if partner_ids else None,
                                        'modified_date': officeModifiedDate if officeModifiedDate else datetime.now(),
                                        'user_id':self.res_user.id
                                    })
                                    self.env.cr.commit()
                                    new_event.append(odooEvent.id)
                                except Exception as e:
                                    self.env['office.import.exceptions'].create({
                                        'last_sync': datetime.now(),
                                        'operation': 'create',
                                        'record_id': event['id'],
                                        'odoo_record_name': event['Subject'],
                                        'description': e,
                                        'skip': True,
                                        'model': 'calender.event',
                                        'endpoint': 'Event',
                                        'sync_type': 'scheduled' if Auto else 'manual',
                                        'user_id': self.res_user.id,
                                    })
                                    self.env.cr.commit()
                                    continue
                            else:
                                if odooEvent[0].write_date:
                                    if str(odooEvent[0].write_date).split('.')[0] >= str(officeModifiedDate):
                                        continue
                                    else:
                                        legth = len(odooEvent)-1
                                        try:
                                            odooEvent[legth].sudo().write({
                                                'office_id': event['id'],
                                                'name': event['subject'],
                                                'calendar_id': self.calendar_id.id if self.calendar_id else None,
                                                'category_name': event['categories'][0] if event['categories'] else None,
                                                "description": event['body']['content'],
                                                'location': location,
                                                'privacy': privacy,
                                                # 'dtstart': start_date if start_date else None,
                                                'until': end_date if end_date else None,
                                                'allday': event['isAllDay'],
                                                'categ_ids': [[6, 0, categories_ids]] if categories_ids else None,
                                                'show_as': event['showAs'],
                                                'recurrency': True if event['recurrence'] else False,
                                                'interval': interval,
                                                'end_type': end_type,
                                                'rrule_type': rRule,
                                                'count': occurences,
                                                'mon': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'monday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'tue': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'tuesday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'wed': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'wednesday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'thu': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'thursday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'fri': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'friday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'sat': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'saturday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'sun': True if event['recurrence'] and 'daysOfWeek' in event['recurrence'][
                                                    'pattern'].keys() and 'sunday' in event['recurrence']['pattern'][
                                                                  'daysOfWeek'] else False,
                                                'month_by': month_by,
                                                'day': day,
                                                'byday': week_index[0] if week_index else None ,
                                                'weekday': week_days[0] if week_days else None ,
                                                'partner_ids': [[6, 0, partner_ids]] if partner_ids else None,
                                                'modified_date': officeModifiedDate if officeModifiedDate else datetime.now(),
                                                'recurrence_update': 'future_events' if odooEvent[0].recurrency else 'self_only',
                                                'user_id': self.res_user.id,
                                            })
                                            self.env.cr.commit()
                                            update_event.append(odooEvent[0].id)
                                        except Exception as e:
                                            self.env['office.import.exceptions'].create({
                                                'last_sync': datetime.now(),
                                                'operation': 'update',
                                                'record_id': event['id'],
                                                'odoo_record_name': event['Subject'],
                                                'description': e,
                                                'skip': True,
                                                'model': 'calender.event',
                                                'endpoint': 'Event',
                                                'sync_type': 'scheduled' if Auto else 'manual',
                                                'user_id': self.res_user.id,
                                            })
                                            self.env.cr.commit()
                                            continue
                            if odoo_event_ids and self.res_user.office365_event_del_flag:
                                delete_event = self.env['calendar.event'].search([('office_id', 'in', odoo_event_ids)])
                                delete_event.unlink()
                        except Exception as e:
                            self.env['office.import.exceptions'].create({
                                'last_sync':datetime.now(),
                                'record_id':event['id'],
                                'odoo_record_name':event['Subject'],
                                'description':e,
                                'skip':True,
                                'model':'calender.event',
                                'endpoint':'Event',
                                'sync_type':'scheduled' if Auto else 'manual',
                                'user_id':self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue
                    if '@odata.nextLink' in responseDecoded:
                        url = responseDecoded['@odata.nextLink']
                    else:
                        break
                else:
                    raise ValidationError(("Unable to connect with office.Kindly build connection again"))
                    break
            return len(new_event), len(update_event),

        except Exception as e:
            traceback.print_exc()
            raise ValidationError(str(e))

    def getEventOdooAttendees(self, attendees):
        try:
            odooAttendees = []
            for attendee in attendees:
                if 'emailAddress' in attendee:
                    partner = self.env['res.partner'].search([('email', "=", attendee['emailAddress']['address'].lower())])
                    if not partner:
                        partner = self.env['res.partner'].create({
                            'name': attendee['emailAddress']['name'],
                            'email': attendee['emailAddress']['address'].lower(),
                            'location': 'Office365 Attendee',
                        })
                        self.env.cr.commit()
                        self.env.cr.execute("""UPDATE res_partner SET create_uid='%s' WHERE id=%s""" % (
                        self.res_user.id, partner.id))
                        self.env.cr.commit()
                        odooAttendees.append(partner.id)
                    else:
                        odooAttendees.append(partner[0].id)

            return odooAttendees
        except Exception as e:
            raise ValidationError(str(e))

    def getAttendee(self, attendees):
        try:
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
        except Exception as e:
            raise ValidationError(str(e))

    def getTime(self, alarm):
        try:
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
        except Exception as e:
            raise ValidationError(str(e))

    def getdays(self, meeting):
        try:
            """
            Returns days of week the event will occure
            :param meeting:
            :return: list of days
            """
            days = []
            if meeting.weekday == "SUN":
                days.append("Sunday")
            if meeting.weekday == "MON":
                days.append("Monday")
            if meeting.weekday == "TUE":
                days.append("Tuesday")
            if meeting.weekday == "WED":
                days.append("Wednesday")
            if meeting.weekday == "THU":
                days.append("Thursday")
            if meeting.weekday == "FRI":
                days.append("Friday")
            if meeting.weekday == "SAT":
                days.append("Saturday")
            return days
        except Exception as e:
            raise ValidationError(str(e))

    def export_calendar(self):
        export_event = []
        update_event = []
        try:
            res_user = self.res_user
            if self.from_date and not self.to_date:
                raise ValidationError('Warning!', 'Please! Select "To Date" to Import Events.')
            if not self.from_date and self.to_date:
                raise ValidationError('Warning!', 'Please! Select "From Date" to Import Events.')
            if self.from_date > self.to_date:
                raise ValidationError('Warning!', 'Please! Enter Date in correct Order !')
            if self.from_date and self.to_date:
                from_date = self.from_date
                to_date = self.to_date

            header = {
                'Authorization': 'Bearer {0}'.format(res_user.token),
                'Content-Type': 'application/json'
            }
            if self.calendar_id and self.calendar_id.calendar_id:
                calendar_id = self.calendar_id.calendar_id
            else:
                response = requests.get(
                    'https://graph.microsoft.com/v1.0/me/calendars',
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(res_user.token),
                        'Accept': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    })
                if response.status_code == 200 or response.status_code == 201:
                    res_data = json.loads((response.content.decode('utf-8')))
                    calendars = res_data['value']
                    calendar_id = calendars[0]['id']
            meetings = self.env['calendar.event'].search([("create_uid", '=', res_user.id)])
            if self.from_date and self.to_date:
                meetings = meetings.search([('write_date', '>=', from_date), ('write_date', '<=', to_date),("create_uid", '=', res_user.id)])

            added = []
            for meeting in meetings:
                temp = meeting
                id = str(meeting.id).split('-')[0]
                metngs = [meeting for meeting in meetings if id in str(meeting.id)]
                index = len(metngs)
                categ_name = []
                if meeting.categ_ids:
                    for cat in meeting.categ_ids:
                        categ_name.append(cat.name)
                if meeting.start and type(meeting.start) is datetime:
                    metting_start = meeting.start.strftime('%Y-%m-%d T %H:%M:%S')
                else:
                    metting_start = meeting.start

                if meeting.stop and type(meeting.stop) is datetime:
                    metting_stop = meeting.stop.strftime('%Y-%m-%d T %H:%M:%S')
                else:
                    metting_stop = meeting.stop
                if meeting.privacy == 'private':
                    privacy = 'private'
                else:
                    privacy = 'normal'
                payload = {
                    "subject": meeting.name,
                    "categories": categ_name,
                    "attendees": self.getAttendee(meeting.attendee_ids),
                    'reminderMinutesBeforeStart': self.getTime(meeting.alarm_ids),
                    "isAllDay":meeting.allday,
                    "body": {
                        'contentType': 'text',
                        'content': meeting.description if meeting.description else None},
                    "sensitivity": privacy,
                    "start": {
                        "dateTime": metting_start,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": metting_stop,
                        "timeZone": "UTC"
                    },
                    "showAs": meeting.show_as,
                    "location": {
                        "displayName": meeting.location if meeting.location else "",
                    },

                }
                if meeting.recurrency:
                    if meeting.recurrence_id.rrule_type == 'daily':
                        if meeting.recurrence_id.end_type == 'count':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'daily',
                                    "interval": meeting.recurrence_id.interval,
                                },
                                "range": {
                                    "type": "numbered",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "numberOfOccurrences" : meeting.recurrence_id.count,
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'end_date':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'daily',
                                    "interval": meeting.recurrence_id.interval,
                                },
                                "range": {
                                    "type": "endDate",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "endDate": str(meeting.recurrence_id.until),
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'forever':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'daily',
                                    "interval": meeting.recurrence_id.interval,
                                },
                                "range": {
                                    "type": "noEnd",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                    if meeting.recurrence_id.rrule_type == 'weekly':
                        weekdays = []
                        days = {'mon': 'Monday', 'tue': 'Tuesday', 'wed': 'Wednesday', 'thu': 'Thursday', 'fri': 'Friday',
                                'sat': 'Saturday', 'sun': 'Sunday'}
                        for day in days:
                            if meeting.recurrence_id[day] == True:
                                weekdays.append(days[day])
                        if meeting.recurrence_id.end_type == 'count':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'weekly',
                                    "interval": meeting.recurrence_id.interval,
                                    "daysOfWeek": weekdays,
                                },
                                "range": {
                                    "type": "numbered",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "numberOfOccurrences": meeting.recurrence_id.count,
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'end_date':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'weekly',
                                    "interval": meeting.recurrence_id.interval,
                                    "daysOfWeek": weekdays,
                                },
                                "range": {
                                    "type": "endDate",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "endDate": str(meeting.recurrence_id.until),
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'forever':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'weekly',
                                    "interval": meeting.recurrence_id.interval,
                                },
                                "range": {
                                    "type": "noEnd",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                    if meeting.recurrence_id.rrule_type == 'monthly':
                        if meeting.recurrence_id.end_type == 'count':
                            if meeting.recurrence_id.month_by == 'date':
                                payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteMonthly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                },
                                "range": {
                                    "type": "numbered",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "numberOfOccurrences": meeting.recurrence_id.count,
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                            if meeting.recurrence_id.month_by == 'day':
                                week_index = 'first'
                                daysOfWeek = self.getdays(meeting)
                                if meeting.byday == '1':
                                    week_index = 'first'
                                elif meeting.byday == '2':
                                    week_index = 'second'
                                elif meeting.byday == '3':
                                    week_index = 'third'
                                elif meeting.byday == '4':
                                    week_index = 'fourth'
                                elif meeting.byday == '-1':
                                    week_index = 'last'
                                payload.update({"recurrence": {
                                    "pattern": {
                                        "type": 'relativeMonthly',
                                        "interval": meeting.recurrence_id.interval,
                                        "daysOfWeek": daysOfWeek,
                                        "index": week_index
                                    },
                                    "range": {
                                        "type": "numbered",
                                        "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                        "numberOfOccurrences": meeting.recurrence_id.count,
                                        "recurrenceTimeZone": "UTC",
                                    }
                                }})
                        if meeting.recurrence_id.end_type == 'end_date':
                            if meeting.recurrence_id.month_by == 'date':
                                payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteMonthly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                },
                                "range": {
                                    "type": "endDate",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "endDate": str(meeting.recurrence_id.until),
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                            if meeting.recurrence_id.month_by == 'day':
                                week_index = 'first'
                                daysOfWeek = self.getdays(meeting)
                                if meeting.byday == '1':
                                    week_index = 'first'
                                elif meeting.byday == '2':
                                    week_index = 'second'
                                elif meeting.byday == '3':
                                    week_index = 'third'
                                elif meeting.byday == '4':
                                    week_index = 'fourth'
                                elif meeting.byday == '-1':
                                    week_index = 'last'
                                payload.update({"recurrence": {
                                    "pattern": {
                                        "type": 'relativeMonthly',
                                        "interval": meeting.recurrence_id.interval,
                                        "daysOfWeek": daysOfWeek,
                                        "index": week_index
                                    },
                                    "range": {
                                    "type": "endDate",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "endDate": str(meeting.recurrence_id.until),
                                    "recurrenceTimeZone": "UTC",
                                }
                                }})
                        if meeting.recurrence_id.end_type == 'forever':
                            if meeting.recurrence_id.month_by == 'date':
                                payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteMonthly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                },
                                "range": {
                                    "type": "noEnd",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                            if meeting.recurrence_id.month_by == 'day':
                                week_index = 'first'
                                daysOfWeek = self.getdays(meeting)
                                if meeting.byday == '1':
                                    week_index = 'first'
                                elif meeting.byday == '2':
                                    week_index = 'second'
                                elif meeting.byday == '3':
                                    week_index = 'third'
                                elif meeting.byday == '4':
                                    week_index = 'fourth'
                                elif meeting.byday == '-1':
                                    week_index = 'last'
                                payload.update({"recurrence": {
                                    "pattern": {
                                        "type": 'relativeMonthly',
                                        "interval": meeting.recurrence_id.interval,
                                        "daysOfWeek": daysOfWeek,
                                        "index": week_index
                                    },
                                    "range": {
                                    "type": "noEnd",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "recurrenceTimeZone": "UTC",
                                }
                                }})
                    if meeting.recurrence_id.rrule_type == 'yearly':
                        if meeting.recurrence_id.end_type == 'count':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteYearly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                    "month": str(meeting.recurrence_id.dtstart).split(" ")[0].split("-")[1],
                                },
                                "range": {
                                    "type": "numbered",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "numberOfOccurrences": meeting.recurrence_id.count,
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'end_date':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteYearly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                    "month": str(meeting.recurrence_id.dtstart).split(" ")[0].split("-")[1],
                                },
                                "range": {
                                    "type": "endDate",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "endDate": str(meeting.recurrence_id.until),
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                        if meeting.recurrence_id.end_type == 'forever':
                            payload.update({"recurrence": {
                                "pattern": {
                                    "type": 'absoluteYearly',
                                    "interval": meeting.recurrence_id.interval,
                                    "dayOfMonth": meeting.recurrence_id.day,
                                    "month": str(meeting.recurrence_id.dtstart).split(" ")[0].split("-")[1],
                                },
                                "range": {
                                    "type": "noEnd",
                                    "startDate": str(meeting.recurrence_id.dtstart).split(" ")[0],
                                    "recurrenceTimeZone": "UTC",
                                }
                            }})
                if meeting.recurrence_id not in added:
                    if not meeting.office_id:
                        newly_create_event_response = requests.post(
                            'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events',
                            headers=header, data=json.dumps(payload))
                        if newly_create_event_response.status_code==200 or newly_create_event_response.status_code==201:
                            if 'id' in json.loads((newly_create_event_response.content.decode('utf-8'))):
                                office_data = json.loads((newly_create_event_response.content.decode('utf-8')))
                                temp.write({
                                    'office_id': office_data['id'],
                                    'modified_date': datetime.strptime(office_data['lastModifiedDateTime'].split(".")[0],'%Y-%m-%dT%H:%M:%S')})
                                temp.is_update = False
                                self.env.cr.commit()
                                export_event.append(json.loads((newly_create_event_response.content.decode('utf-8')))['id'])
                                if meeting.recurrency:
                                    added.append(meeting.recurrence_id)
                        else:
                            newly_create_response_exception = json.loads(newly_create_event_response.content.decode('utf-8'))
                            self.env['office.export.exceptions'].create({
                                'last_sync': datetime.now(),
                                'endpoint': 'Calender',
                                'operation': 'create',
                                'contact_id': None,
                                'event_id': meeting.id,
                                'task_id': None,
                                'skip': True,
                                'description': newly_create_response_exception['error']['code'] + '   ' + newly_create_response_exception['error'][
                                    'message'],
                                'model': 'Event',
                                # 'sync_type': 'scheduled' if Auto else 'manual',
                                'user_id': self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue
                    else:
                        if meeting.office_id:
                            url = "https://graph.microsoft.com/v1.0/me/events/{}".format(meeting.office_id)
                            redundency_check_response = requests.get(
                                url,
                                headers={
                                    'Host': 'outlook.office.com',
                                    'Authorization': 'Bearer {0}'.format(res_user.token),
                                    'Accept': 'application/json',
                                    'X-Target-URL': 'http://outlook.office.com',
                                    'connection': 'keep-Alive'
                                })
                            if redundency_check_response.status_code == 200 or redundency_check_response.status_code == 201:
                                res_event = json.loads((redundency_check_response.content.decode('utf-8')))
                                modified_at = datetime.strptime(
                                    res_event['lastModifiedDateTime'].split(".")[0], '%Y-%m-%dT%H:%M:%S'
                                )
                                # if modified_at != meeting.modified_date:
                                if meeting.write_date:
                                    if str(modified_at)>str(meeting.write_date).split('.')[0] or str(modified_at) == str(meeting.write_date).split('.')[0]:
                                        continue

                                update_event_response = requests.patch(
                                    'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events/' + meeting.office_id,
                                    headers=header, data=json.dumps(payload))
                                if update_event_response.status_code == 200 or update_event_response.status_code == 201:
                                    if 'id' in json.loads((update_event_response.content.decode('utf-8'))):
                                        temp.write({
                                            'office_id': json.loads((update_event_response.content.decode('utf-8')))['id'],
                                            'modified_date': modified_at
                                        })
                                        update_event.append(json.loads((update_event_response.content.decode('utf-8')))['id'])
                                        meeting.is_update = False
                                        self.env.cr.commit()
                                        if meeting.recurrency:
                                            added.append(id)
                                else:
                                    update_event_response_exception = json.loads(update_event_response.content.decode('utf-8'))
                                    self.env['office.export.exceptions'].create({
                                        'last_sync': datetime.now(),
                                        'endpoint': 'Calender',
                                        'operation': 'update',
                                        'contact_id': None,
                                        'event_id': meeting.id,
                                        'task_id': None,
                                        'skip': True,
                                        'description': update_event_response_exception['error']['code'] + ' ' +
                                                       update_event_response_exception['error'][
                                                           'message'],
                                        'model': 'Event',
                                        # 'sync_type': 'scheduled' if Auto else 'manual',
                                        'user_id': self.res_user.id,
                                    })
                                    self.env.cr.commit()
                                    continue
                            else:
                                not_found_calender_create = requests.post(
                                    'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events',
                                    headers=header, data=json.dumps(payload))
                                if not_found_calender_create.status_code == 200 or not_found_calender_create.status_code == 201:
                                    if 'id' in json.loads((not_found_calender_create.content.decode('utf-8'))):
                                        office_data=json.loads((not_found_calender_create.content.decode('utf-8')))
                                        temp.write({
                                            'office_id': office_data['id'],
                                            'modified_date': datetime.strptime(office_data['lastModifiedDateTime'].split(".")[0], '%Y-%m-%dT%H:%M:%S')
                                        })
                                        temp.is_update = False
                                        self.env.cr.commit()
                                        export_event.append(
                                            json.loads((not_found_calender_create.content.decode('utf-8')))['id'])
                                        if meeting.recurrency:
                                            added.append(id)
                                else:
                                    not_found_calender_create_exception = json.loads(not_found_calender_create.content.decode('utf-8'))
                                    self.env['office.export.exceptions'].create({
                                        'last_sync': datetime.now(),
                                        'endpoint': 'Calender',
                                        'operation': 'create',
                                        'contact_id': None,
                                        'event_id': meeting.id,
                                        'task_id': None,
                                        'skip': True,
                                        'description': not_found_calender_create_exception['error']['code'] + ' ' +
                                                       not_found_calender_create_exception['error'][
                                                           'message'],
                                        'model': 'Event',
                                        # 'sync_type': 'scheduled' if Auto else 'manual',
                                        'user_id': self.res_user.id,
                                    })
                                    self.env.cr.commit()
                                    continue
                        else:
                            new_event_create = requests.post(
                                'https://graph.microsoft.com/v1.0/me/calendars/' + calendar_id + '/events',
                                headers=header, data=json.dumps(payload))
                            if new_event_create.status_code == 200 or new_event_create.status_code == 201:
                                if 'id' in json.loads((new_event_create.content.decode('utf-8'))):
                                    office_data = json.loads((new_event_create.content.decode('utf-8')))
                                    temp.write({
                                        'office_id': office_data['id'],
                                        'modified_date': datetime.strptime(office_data['lastModifiedDateTime'].split(".")[0], '%Y-%m-%dT%H:%M:%S')})
                                    temp.is_update = False
                                    self.env.cr.commit()
                                    export_event.append(json.loads((new_event_create.content.decode('utf-8')))['id'])
                                    if meeting.recurrency:
                                        added.append(id)
                            else:
                                new_event_create_exception = json.loads(new_event_create.content.decode('utf-8'))
                                self.env['office.export.exceptions'].create({
                                    'last_sync': datetime.now(),
                                    'endpoint': 'Calender',
                                    'operation': 'create',
                                    'contact_id': None,
                                    'event_id': meeting.id,
                                    'task_id': None,
                                    'skip': True,
                                    'description': new_event_create_exception['error']['code'] + ' ' +
                                                   new_event_create_exception['error'][
                                                       'message'],
                                    'model': 'Event',
                                    # 'sync_type': 'scheduled' if Auto else 'manual',
                                    'user_id': self.res_user.id,
                                })
                                self.env.cr.commit()
                                continue

            export_dictionary = {
            'exportedCalenders': len(export_event),
            'updatedCalenders': len(update_event)
            }

            return export_dictionary

        except Exception as e:
            _logger.error(e)
            raise ValidationError(_(str(e)))


    '''
        These following methods are responsible fro importing tasks from Office 365
    '''

    def import_tasks(self):
        try:
            office_tasks = 0
            update_tasks = 0
            odooUser = self.res_user
            url = 'https://graph.microsoft.com/v1.0/me/todo/lists'
            header = {
                'Authorization': 'Bearer {0}'.format(odooUser.token),
                'Content-type': 'application/json',
                'connection': 'keep-Alive'
            }
            response = requests.get(url, headers=header)
            if response.status_code == 200:
                res_data = json.loads((response.content.decode('utf-8')))
                for list in res_data['value']:
                    new_url = 'https://graph.microsoft.com/v1.0/me/todo/lists/' + list['id'] + '/tasks'

                    if not self.from_date and not self.to_date:
                        office_tasks, update_tasks = self.create_tasks(new_url)
                    if self.from_date and self.to_date:
                        url = new_url + '?$filter=lastModifiedDateTime ge {} and lastModifiedDateTime le {}' \
                                  .format(self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                                          self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                        office_tasks, update_tasks = self.create_tasks(url)

            import_dictionary = {
                'importedTasks': office_tasks,
                'updatedTasks': update_tasks,
            }

            return import_dictionary
        except Exception as e:
            raise ValidationError(str(e))

    def create_tasks(self, url):
        try:
            new_tasks = []
            update_tasks = []
            headers = {
                'Authorization': 'Bearer {0}'.format(self.res_user.token),
                'Accept': 'application/json',
                'connection': 'keep-Alive'}
            while True:
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    res_data = json.loads((response.content.decode('utf-8')))
                    tasks = res_data['value']
                    for task in tasks:
                        try:
                            interval_selected=[]
                            res_model_id = self.env.ref('odoo_office365.model_activity_general')
                            res_id = self.env.ref('odoo_office365.general_activities')
                            odoo_task = self.env['mail.activity'].sudo().search([('office_id', '=', task['id'])])
                            dueDateTime = task['dueDateTime']['dateTime'][:-16] if 'dueDateTime' in task else None
                            officeModifiedDate = datetime.strptime(datetime.strftime(p.parse(task['lastModifiedDateTime']), "%Y-%m-%dT%H:%M:%S"),"%Y-%m-%dT%H:%M:%S")
                            activity_type = self.env['mail.activity.type'].search([('name', '=', 'To Do')])
                            intervals = {'daily': 'Daily', 'weekdays': 'Weekdays', 'weekly': 'Weekly','absoluteMonthly': 'Monthly', 'absoluteYearly': 'Yearly'}
                            if 'recurrence' in task:
                                for interval in intervals:
                                    if task['recurrence']['pattern']['type'] == interval:
                                        if interval == 'weekly' and task['recurrence']['pattern']['daysOfWeek']:
                                            interval_selected.append(intervals['weekdays'])
                                        else:
                                            interval_selected.append(intervals[interval])
                            if not odoo_task:
                                try:
                                    odooTask = self.env['mail.activity'].create({
                                        'res_id': res_id.id,
                                        'activity_type_id': activity_type.id if activity_type else None,
                                        'summary': task['title'],
                                        'date_deadline': (datetime.strptime(dueDateTime, '%Y-%m-%dT')).strftime('%Y-%m-%d') if dueDateTime else datetime.now(),
                                        'note': task['body']['content'].replace('\n', '<br>'),
                                        'res_model_id': res_model_id.id,
                                        'recurring': True if 'recurrence' in task else False,
                                        'interval':interval_selected[0] if interval_selected else None,
                                        'office_id': task['id'],
                                        'user_id': self.res_user.id,
                                        'modified_date': datetime.strptime(datetime.strftime(p.parse(task['lastModifiedDateTime']), "%Y-%m-%dT%H:%M:%S"),"%Y-%m-%dT%H:%M:%S"),
                                    })
                                    self.env.cr.commit()
                                    self.env.cr.execute("""UPDATE mail_activity SET create_uid='%s' WHERE id=%s""" % (self.res_user.id, odooTask.id))
                                    self.env.cr.commit()
                                    if 'recurrence' in task:
                                        new_date=odooTask.get_date()
                                        odooTask.write({
                                            'new_date':new_date,
                                        })
                                    if task['status'] == 'completed':
                                        odooTask.action_done()
                                    new_tasks.append(odooTask.id)
                                except Exception as e:
                                    self.env['office.import.exceptions'].create({
                                        'last_sync': datetime.now(),
                                        'record_id': task['id'],
                                        'odoo_record_name': task['title'],
                                        'description': e,
                                        'skip': True,
                                        'model': 'mail.activity',
                                        'endpoint': 'Tasks',
                                        # 'sync_type':'scheduled' if Auto else 'manual',
                                        'user_id': self.res_user.id,
                                    })
                                    self.env.cr.commit()
                                    continue
                            else:
                                try:
                                    if odoo_task.modified_date:
                                        if odoo_task.modified_date >= officeModifiedDate:
                                            continue
                                        else:
                                            odoo_task.write({
                                                'res_id': res_id.id,
                                                'activity_type_id': activity_type.id if activity_type else None,
                                                'summary': task['title'],
                                                'date_deadline': (datetime.strptime(dueDateTime, '%Y-%m-%dT')).strftime('%Y-%m-%d')
                                                if dueDateTime else datetime.now(),
                                                'note': task['body']['content'].replace('\n', '<br>'),
                                                'res_model_id': res_model_id.id,
                                                'office_id': task['id'],
                                                'user_id': self.res_user.id,
                                                'modified_date': datetime.strptime(datetime.strftime(p.parse(
                                                    task['lastModifiedDateTime']), "%Y-%m-%dT%H:%M:%S"),
                                                    "%Y-%m-%dT%H:%M:%S"),
                                            })
                                            if task['status'] == 'completed':
                                                odoo_task.action_done()
                                            update_tasks.append(odoo_task.id)
                                    self.env.cr.commit()
                                except Exception as e:
                                    self.env['office.import.exceptions'].create({
                                        'last_sync': datetime.now(),
                                        'record_id': task['id'],
                                        'odoo_record_name': task['title'],
                                        'description': e,
                                        'skip': True,
                                        'model': 'mail.activity',
                                        'endpoint': 'Tasks',
                                        # 'sync_type':'scheduled' if Auto else 'manual',
                                        'user_id': self.res_user.id,
                                    })
                                    self.env.cr.commit()
                                    continue
                        except Exception as e:
                            self.env['office.import.exceptions'].create({
                                'last_sync':datetime.now(),
                                'record_id':task['id'],
                                'odoo_record_name':task['title'],
                                'description':e,
                                'skip':True,
                                'model':'mail.activity',
                                'endpoint':'Tasks',
                                # 'sync_type':'scheduled' if Auto else 'manual',
                                'user_id':self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue
                    if '@odata.nextLink' in res_data:
                        url = res_data['@odata.nextLink']
                    else:
                        break
                else:
                    raise ValidationError(("Unable to connect with office.Kindly build connection again"))
                    break
            return len(new_tasks), len(update_tasks)
        except Exception as e:
            raise ValidationError(_(str(e)))

    def export_tasks(self):
        export_task = []
        update_task = []
        task_list_id = None
        try:
            res_user = self.res_user
            if self.from_date and not self.to_date:
                raise ValidationError('Warning!', 'Please! Select "To Date" to Import Events.')
            if not self.from_date and self.to_date:
                raise ValidationError('Warning!', 'Please! Select "From Date" to Import Events.')
            if self.from_date > self.to_date:
                raise ValidationError('Warning!', 'Please! Enter Date in correct Order !')
            if self.from_date and self.to_date:
                from_date = self.from_date
                to_date = self.to_date
            odoo_activities = self.env['mail.activity'].search([('user_id', '=', self.res_user.id)])
            if self.from_date and self.to_date:
                odoo_activities = odoo_activities.search(
                    [('write_date', '>=', self.from_date), ('write_date', '<=', self.to_date)])
            task_list_id_response = requests.get(url='https://graph.microsoft.com/v1.0/me/todo/lists/tasks', headers={
                'Authorization': 'Bearer {0}'.format(res_user.token),
                'Content-type': 'application/json',
                'connection': 'keep-Alive'
            })
            if task_list_id_response.status_code == 200 or task_list_id_response.status_code == 201:
                task_list_id_json = json.loads((task_list_id_response.content.decode('utf-8')))
                if 'id' in task_list_id_json:
                    task_list_id = task_list_id_json['id']
            for activity in odoo_activities:
                headers = {
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(res_user.token),
                    'Accept': 'application/json',
                    'Content-type': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }
                payload = {
                    'title': activity.summary if activity.summary else activity.display_name,
                    "body": {
                        "contentType": "html",
                        "content": activity.note if activity.note else ""
                    },
                    "dueDateTime": {
                        "dateTime": str(activity.date_deadline) + 'T00:00:00Z',
                        "timeZone": "UTC"
                    },
                }
                if activity.recurring:
                    intervals = {'Daily': 'daily', 'Weekdays': 'weekly', 'Weekly': 'weekly',
                                 'Monthly': 'absoluteMonthly', 'Yearly': 'absoluteYearly'}
                    for interval in intervals:
                        if activity.interval == interval:
                            if activity.interval == 'Weekdays':
                                payload.update({"recurrence": {
                                    "pattern": {
                                        "type": intervals.get(interval),
                                        "daysOfWeek": ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']
                                    }}})
                            else:
                                payload.update({"recurrence": {
                                    "pattern": {
                                        "type": intervals.get(interval),
                                    }}})
                if not activity.office_id:
                    task_url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + task_list_id + "/tasks"
                    newly_create_task_response = requests.post(url=task_url, data=json.dumps(payload), headers=headers)
                    if newly_create_task_response.status_code == 200 or newly_create_task_response.status_code == 201:
                        if 'id' in json.loads((newly_create_task_response.content.decode('utf-8'))).keys():
                            activity.office_id = json.loads((newly_create_task_response.content.decode('utf-8')))['id']
                            export_task.append(activity.office_id)
                            activity.is_update = False
                            self.env.cr.commit()
                    else:
                        new_task_create_exception = json.loads(newly_create_task_response.content.decode('utf-8'))
                        self.env['office.export.exceptions'].create({
                            'last_sync': datetime.now(),
                            'endpoint': 'Tasks',
                            'operation': 'create',
                            'contact_id': None,
                            'event_id': None,
                            'task_id': activity.id,
                            'skip': True,
                            'description': new_task_create_exception['error']['code'] + ' ' +
                                           new_task_create_exception['error']['message'],
                            'model': 'mail.activity',
                            # 'sync_type': 'scheduled' if Auto else 'manual',
                            'user_id': self.res_user.id,
                        })
                        self.env.cr.commit()
                        continue
                else:
                    task_url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + task_list_id + "/tasks/{}".format(
                        activity.office_id)
                    redundency_check_response = requests.get(task_url, headers={
                        'Authorization': 'Bearer {0}'.format(res_user.token),
                        'Content-type': 'application/json',
                        'connection': 'keep-Alive'
                    })
                    if redundency_check_response.status_code == 200 or redundency_check_response.status_code == 201:
                        res_tasks = json.loads((redundency_check_response.content.decode('utf-8')))
                        modified_at = datetime.strptime(res_tasks['lastModifiedDateTime'].split(".")[0],
                                                        '%Y-%m-%dT%H:%M:%S')
                        if activity.write_date:
                            if str(modified_at) > str(activity.write_date).split('.')[0] or str(modified_at) == \
                                    str(activity.write_date).split('.')[0]:
                                continue
                        update_task_response = requests.patch(
                            url="https://graph.microsoft.com/v1.0/me/todo/lists/" + task_list_id + "/tasks/{}".format(
                                activity.office_id), headers=headers, data=json.dumps(payload))
                        if update_task_response.status_code == 200 or update_task_response.status_code == 201:
                            if 'id' in json.loads((update_task_response.content.decode('utf-8'))):
                                activity.write({
                                    'office_id': json.loads((update_task_response.content.decode('utf-8')))['id'],
                                })
                                update_task.append(json.loads((update_task_response.content.decode('utf-8')))['id'])
                                self.env.cr.commit()
                        else:
                            new_task_update_exception = json.loads(update_task_response.content.decode('utf-8'))
                            self.env['office.export.exceptions'].create({
                                'last_sync': datetime.now(),
                                'endpoint': 'Tasks',
                                'operation': 'create',
                                'contact_id': None,
                                'event_id': None,
                                'task_id': activity.id,
                                'skip': True,
                                'description': new_task_update_exception['error']['code'] + ' ' +
                                               new_task_update_exception['error']['message'],
                                'model': 'mail.activity',
                                # 'sync_type': 'scheduled' if Auto else 'manual',
                                'user_id': self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue
                    else:
                        task_url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + task_list_id + "/tasks"
                        newly_create_task_response = requests.post(url=task_url, data=json.dumps(payload),headers=headers)
                        if newly_create_task_response.status_code == 200 or newly_create_task_response.status_code == 201:
                            if 'id' in json.loads((newly_create_task_response.content.decode('utf-8'))).keys():
                                activity.office_id = json.loads((newly_create_task_response.content.decode('utf-8')))[
                                    'id']
                                export_task.append(activity.office_id)
                                activity.is_update = False
                                self.env.cr.commit()
                        else:
                            new_task_create_exception = json.loads(newly_create_task_response.content.decode('utf-8'))
                            self.env['office.export.exceptions'].create({
                                'last_sync': datetime.now(),
                                'endpoint': 'Tasks',
                                'operation': 'create',
                                'contact_id': None,
                                'event_id': None,
                                'task_id': activity.id,
                                'skip': True,
                                'description': new_task_create_exception['error']['code'] + ' ' +
                                               new_task_create_exception['error']['message'],
                                'model': 'mail.activity',
                                # 'sync_type': 'scheduled' if Auto else 'manual',
                                'user_id': self.res_user.id,
                            })
                            self.env.cr.commit()
                            continue

            export_dictionary = {
                'exportedTasks': len(export_task),
                'updatedTasks': len(update_task)
            }
            return export_dictionary

        except Exception as e:
            raise ValidationError(_(str(e)))

    '''
        These following methods are responsible fro importing emails from Office 365
    '''

    def sync_customer_mail(self):
        try:
            office_emails =0
            url = 'https://graph.microsoft.com/v1.0/me/mailFolders'
            headers = {
                'Authorization': 'Bearer {0}'.format(self.res_user.token),
                'Content-type': 'application/json',
                'connection': 'keep-Alive'
            }
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                folders = json.loads((response.content.decode('utf-8')))['value']
                folder_ids = [folder['id'] for folder in folders if folder['displayName'] == 'Sent Items' or folder['displayName'] == 'Inbox']
                for folder_id in folder_ids:
                    if self.from_date and self.to_date:
                        url = 'https://graph.microsoft.com/v1.0/me/mailFolders/' + folder_id + \
                              '/messages?$top=1000&$count=true&$filter=ReceivedDateTime ge {} and ReceivedDateTime le {}' \
                                  .format(self.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                                          self.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                        office_emails = self.create_emails(url)
                    if not self.from_date and not self.to_date:
                        url = 'https://graph.microsoft.com/v1.0/me/mailFolders/' + folder_id + '/messages?$top=1000&$count=true'
                        office_emails = self.create_emails(url)

            import_dictionary = {
                'importedEmails': office_emails,
                }

            return import_dictionary
        except Exception as e:
            raise ValidationError(str(e))

    def create_emails(self, url):
        try:
            new_email = []

            headers = {
                'Authorization': 'Bearer {0}'.format(self.res_user.token),
                'Content-type': 'application/json',
                'connection': 'keep-Alive'
            }
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                messages = json.loads((response.content))
                for message in messages['value']:
                    attachment_ids = []
                    odooMail = self.env['mail.message'].search([('office_id', '=', message['id'])])
                    if not odooMail:
                        if self.checkrequired(message) == 0:
                            continue
                        if message['hasAttachments']:
                            attachment_ids = self.getAttachment(message)
                        from_partner = self.env['res.partner'].search([('email', "=", message['from']['emailAddress']['address'])])
                        if not from_partner:
                            continue
                        else:
                            from_partner = from_partner[0]
                        recipient_partners = []
                        for recipient in message['toRecipients']:
                            odooUser = self.env['res.users'].search([('email', '=', recipient['emailAddress']['address'].lower())])
                            if odooUser:
                                odooPartner = odooUser[0].partner_id
                                recipient_partners.append(odooPartner.id)

                        date = datetime.strptime(message['sentDateTime'], "%Y-%m-%dT%H:%M:%SZ")
                        newMail = self.env['mail.message'].create({
                            'subject': message['subject'],
                            'date': date,
                            'body': message['body']['content'] if 'content' in message['body'] else '',
                            'email_from': message['sender']['emailAddress']['address'],
                            'partner_ids': [[6, 0, recipient_partners]],
                            'attachment_ids': [[6, 0, attachment_ids]],
                            'office_id': message['id'],
                            'author_id': from_partner.id,
                            'model': 'res.partner',
                            'res_id': from_partner.id,
                        })
                        self.env.cr.commit()

                        new_email.append(newMail.id)
                return len(new_email)
            else:
                raise ValidationError(("Unable to connect with office Mail.Kindly build connection again"))
        except Exception as e:
            raise ValidationError(str(e))

    ''' 
                    This following method is responsible for extracting Webhook Message and pass it to message_new methoed of 'ibs.ticket' module
    '''

    def extract_webhook_message(self, webhook_data):
        try:
            odoo_user = None
            odoo_webhook_subscription = self.env['office.webhook'].sudo().search(
                [('subscription_id', '=', webhook_data['subscriptionId'])])
            if odoo_webhook_subscription:
                odoo_user = odoo_webhook_subscription.user_id
            if odoo_user:
                if odoo_user.token:
                    self.checkTokenExpiryDate(odoo_user)
                    url = 'https://graph.microsoft.com/v1.0/' + webhook_data['resource']
                    headers = {
                        'Authorization': 'Bearer {0}'.format(odoo_user.token),
                        'Content-type': 'application/json',
                        'connection': 'keep-Alive'
                    }
                    response = requests.get(url, headers=headers)
                    if response.status_code == 200:
                        message = json.loads((response.content))
                        from_partner = self.env['res.partner'].sudo().search(
                            [('email', "=", message['from']['emailAddress']['address'])])
                        if not from_partner:
                            from_partner = self.env['res.partner'].sudo().create({
                                'name': message['from']['emailAddress']['address'].split('@')[0],
                                'email': message['from']['emailAddress']['address']
                            })
                            self.env.cr.commit()
                        else:
                            from_partner = from_partner[0]
                        odoo_leads = self.env['crm.lead'].sudo().search([('office_id', '=', message['id'])])
                        if not odoo_leads:
                            self.env['crm.lead'].sudo().create({
                                'name': message['subject'] if message['subject'] else ' ',
                                'partner_id': from_partner.id,
                                'description': message['body']['content'] if 'content' in message['body'] else '',
                            })
                            self.env.cr.commit()
                        attachment_ids = []
                        odoo_mail = self.env['mail.message'].sudo().search([('office_id', '=', message['id'])])
                        if not odoo_mail:
                            if self.checkrequired(message) == 0:
                                return
                            if message['hasAttachments']:
                                attachment_ids = self.getAttachment(message,odoo_user)
                            recipient_partners = []
                            for recipient in message['toRecipients']:
                                odooUser = self.env['res.users'].sudo().search([('email', '=', recipient['emailAddress']['address'].lower())])
                                if odooUser:
                                    odooPartner = odooUser[0].partner_id
                                    recipient_partners.append(odooPartner.id)
                                else:
                                    odooPartner = odoo_user.partner_id
                                    if not odooPartner.id in recipient_partners:
                                        recipient_partners.append(odooPartner.id)

                            date = datetime.strptime(message['sentDateTime'], "%Y-%m-%dT%H:%M:%SZ")
                            newMail = self.env['mail.message'].sudo().create({
                                'subject': message['subject'],
                                'date': date,
                                'body': message['body']['content'] if 'content' in message['body'] else '',
                                'email_from': message['sender']['emailAddress']['address'],
                                'partner_ids': [[6, 0, recipient_partners]],
                                'attachment_ids': [[6, 0, attachment_ids]],
                                'office_id': message['id'],
                                'author_id': from_partner.id,
                                'model': 'res.partner',
                                'res_id': from_partner.id,
                            })
                            self.env.cr.commit()



        except Exception as e:
            raise ValidationError(str(e))

    def getAttachment(self, message,odoo_user_webhook=None):
        try:
            if odoo_user_webhook:
                odoo_user=odoo_user_webhook
            else:
                odoo_user = self.res_user
            url = 'https://graph.microsoft.com/v1.0/me/messages/' + message['id'] + '/attachments/'
            header = {
                    'Host': 'outlook.office.com',
                    'Authorization': 'Bearer {0}'.format(odoo_user.token),
                    'Accept': 'application/json',
                    'X-Target-URL': 'http://outlook.office.com',
                    'connection': 'keep-Alive'
                }

            response = requests.get(url, headers=header)
            if response.status_code == 200:
                attachments = json.loads((response.content))
                attachment_ids = []
                for attachment in attachments['value']:
                    if 'contentBytes' not in attachment or 'name' not in attachment:
                        continue
                    bytes = attachment['contentBytes'].encode("utf-8")
                    odoo_attachment = self.env['ir.attachment'].sudo().create({
                        'datas': bytes,
                        'name': attachment["name"],
                        'res_model': 'mail.message',
                        'store_fname': attachment["name"]
                    })
                    self.env.cr.commit()
                    attachment_ids.append(odoo_attachment.id)
                return attachment_ids
        except Exception as e:
            raise ValueError(str(e))

    def checkrequired(self, message):
        try:
            if 'from' not in message.keys():
                return 0

            elif 'address' not in message.get('from').get('emailAddress') or message['bodyPreview'] == "":
                return 0
            else:
                return 1
        except Exception as e:
            raise ValidationError(str(e))

    '''
        These following methods are responsible for managing token and refresh token
    '''

    def checkTokenExpiryDate(self, odooUser):
        try:
            if odooUser.expires_in:
                expires_in = datetime.fromtimestamp(int(odooUser.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                now = datetime.now()
                if now > expires_in:
                    self.generate_refresh_token(odooUser)
        except Exception as e:
            raise ValidationError(str(e))

    def generate_refresh_token(self, odooUser):
        try:
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

    ''' 
        These following methods are responsible for managing contact's and calendar's categories
    '''

    def getContactsOdooCategory(self, categories):
        try:
            categ_ids = []
            for category in categories:
                odooCategory = self.env['res.partner.category'].search([('name', '=', category)])
                if odooCategory:
                    categ_ids.append(odooCategory[0].id)
                else:
                    newCategory = self.env['res.partner.category'].create({'name': category})
                    categ_ids.append(newCategory.id)
            return categ_ids
        except Exception as e:
            raise ValidationError(str(e))

    def getEventOdooCategory(self, event):
        try:
            categ_id = []
            for categ in event['categories']:
                categ_type_id = self.env['calendar.event.type'].search([('name', '=', categ)])
                if categ_type_id:
                    categ_type_id.write({'name': categ})
                    categ_id.append(categ_type_id[0].id)
                else:
                    categ_type_id = categ_type_id.create({'name': categ})
                    categ_id.append(categ_type_id[0].id)
            return categ_id
        except Exception as e:
            raise ValidationError(str(e))



