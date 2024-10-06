from odoo import models, fields, api, _


class Office_import_Logs(models.Model):
    _name = 'office.export.exceptions'
    # _order = 'id desc'

    endpoint = fields.Char('Office Endpoint')
    operation = fields.Selection(string="Sync Type", selection=[('create', 'Create'), ('update', 'Update')])
    last_sync = fields.Datetime(string="Last Sync", required=False)
    contact_id = fields.Many2one(comodel_name='res.partner', string='Odoo Contact Rec')
    event_id = fields.Many2one(comodel_name='calendar.event', string='Odoo Event Rec')
    task_id = fields.Many2one(comodel_name='mail.activity', string='Odoo Task Rec')
    description = fields.Char('Description')
    skip = fields.Boolean('Skip?')
    model = fields.Char('Odoo Model')
    sync_type = fields.Selection(string="Sync Type", selection=[('scheduled', 'Scheduled'), ('manual', 'Manual')],)
    user_id = fields.Many2one('res.users', default=lambda self: self.env.user)


