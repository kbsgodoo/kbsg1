from datetime import timedelta, datetime
from odoo import models, fields, api
from odoo.tools import DEFAULT_SERVER_DATE_FORMAT
from dateutil.relativedelta import relativedelta

class MailActivity(models.Model):
    _name = 'mail.activity'
    _inherit = ['mail.activity', 'mail.thread']
    _rec_name = 'summary'

    user_id = fields.Many2one('res.users', string='user', index=True,
                              tracking=True, default=lambda self: self.env.user)

    res_model_id = fields.Many2one('ir.model', 'Document Model', index=True, ondelete='cascade', required=True,
                                   default=lambda self: self.env.ref('odoo_office365.model_activity_general'))
    res_id = fields.Many2oneReference(string='Related Document ID', index=True, required=True, model_field='res_model',
                                      default=lambda self: self.env.ref('odoo_office365.general_activities'))

    priority = fields.Selection([
        ('0', 'Normal'),
        ('1', 'Important'),
        ('2', 'Very Important'),
        ('3', 'Urgent'),
    ], default='0', index=True, store=True)
    recurring = fields.Boolean(string="Recurring", store=True)
    state = fields.Selection([
        ('today', 'Today'),
        ('planned', 'Planned'),
        ('done', 'Done'),
        ('overdue', 'Expired'),
        ('cancel', 'Cancelled'), ], 'State',
        compute='_compute_state', store=True)
    interval = fields.Selection(
        [('Daily', 'Daily'),
         ('Weekdays', 'Weekdays'),
         ('Weekly', 'Weekly'),
         ('Monthly', 'Monthly'),
         ('Yearly', 'Yearly')],
        string='Recurring Interval', )
    new_date = fields.Date(string="Next Due Date", store=True)

    def action_done(self):
        """Function done button"""
        self.write({'state': 'done'})
        if self.recurring:
            self.env['mail.activity'].create({
                'res_id': self.res_id,
                'res_model_id': self.res_model_id.id,
                'summary': self.summary,
                'priority': self.priority,
                'date_deadline': self.new_date,
                'recurring': self.recurring,
                'interval': self.interval,
                'activity_type_id': self.activity_type_id.id,
                'new_date': self.get_date(),
                'user_id': self.user_id.id
            })

    def action_cancel(self):
        """ function for cancel button"""
        return self.write({'state': 'cancel'})

    def get_date(self):
        """ function for get new due date on new record"""
        date_deadline = self.new_date if self.new_date else self.date_deadline
        new_date = False
        if self.interval == 'Daily':
            new_date = (
                    date_deadline + timedelta(days=1)).strftime(
                DEFAULT_SERVER_DATE_FORMAT)
        elif self.interval == 'Weekdays':
            dayCheck = date_deadline.weekday()
            if dayCheck == 0 or dayCheck == 1 or dayCheck == 2 or dayCheck == 3:
                new_date = (
                        date_deadline + timedelta(days=1)).strftime(
                    DEFAULT_SERVER_DATE_FORMAT)
            elif dayCheck == 4:
                new_date = (date_deadline + timedelta(days=3)).strftime(DEFAULT_SERVER_DATE_FORMAT)
        elif self.interval == 'Weekly':
            new_date = (date_deadline + timedelta(days=7)).strftime(DEFAULT_SERVER_DATE_FORMAT)
        elif self.interval == 'Monthly':
            new_date = (date_deadline + relativedelta(months=1)).strftime(DEFAULT_SERVER_DATE_FORMAT)
        elif self.interval == 'Yearly':
            # new_date = (date_deadline + timedelta(days=365)).strftime(DEFAULT_SERVER_DATE_FORMAT)
            new_date = (date_deadline + relativedelta(years=1)).strftime(DEFAULT_SERVER_DATE_FORMAT)
        return new_date

    @api.onchange('interval', 'date_deadline')
    def onchange_recurring(self):
        """ function for show new due date"""
        self.new_date = False
        if self.recurring:
            self.new_date = self.get_date()

    def action_date(self):
        """ Function for automated actions for deadline"""
        today = fields.date.today()
        dates = self.env['mail.activity'].search(
            [('state', 'in', ['today', 'planned']),
             ('date_deadline', '=', today),
             ('recurring', '=', True)])
        for rec in dates:
            self.env['mail.activity'].create(
                {'res_id': rec.res_id,
                 'res_model_id': rec.res_model_id.id,
                 'summary': rec.summary,
                 'priority': rec.priority,
                 'interval': rec.interval,
                 'recurring': rec.recurring,
                 'date_deadline': rec.new_date,
                 'new_date': rec.get_date(),
                 'activity_type_id': rec.activity_type_id.id,
                 'user_id': rec.user_id.id
                 })
            rec.state = 'done'


class Office365Activities(models.Model):
    _description = 'Office 365 Activities'
    _name = 'activity.general'
    _inherit = ['mail.thread', 'mail.activity.mixin']

    name = fields.Char('Name')
