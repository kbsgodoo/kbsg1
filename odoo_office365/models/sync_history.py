import logging
import re
from odoo.exceptions import ValidationError, UserError
from odoo.osv import osv
from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools


class HistoryLine(models.Model):
    _name = 'sync.history'
    _description = "Import History"
    _order = 'last_sync desc'

    user_id = fields.Many2one('res.users', default=lambda self: self.env.user)
    last_sync = fields.Datetime(string="Last Sync", required=False)
    numberContacts = fields.Integer(string="New Contacts", required=False)
    numberContactsUpdated = fields.Integer(string="Updated Contacts", required=False)
    numberEmails = fields.Integer(string="New Emails", required=False)
    numberTasks = fields.Integer(string="New Tasks", required=False)
    numberTasksUpdated = fields.Integer(string="Updated Tasks", required=False)
    numberCalendarsUpdated = fields.Integer(string="Updated Events", required=False)
    numberCalendars = fields.Integer(string="New Events", required=False)
    status = fields.Selection([('import', 'Import'), ('export', 'Export')], 'Status')
    sync_type = fields.Selection(string="Sync Type", selection=[('scheduled', 'Scheduled'), ('manual', 'Manual')],required=False)
