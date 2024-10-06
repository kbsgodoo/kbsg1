from odoo import models, fields, api, _


class Office_import_Logs(models.Model):
    _name = 'office.import.exceptions'
    _order = 'id desc'

    record_id = fields.Char('Office Record Id')
    operation = fields.Selection(string="Sync Type", selection=[('create', 'Create'), ('update', 'Update')])
    last_sync = fields.Datetime(string="Last Sync", required=False)
    endpoint = fields.Char('Office Endpoint')
    odoo_record_name = fields.Char('Odoo Record Name')
    description = fields.Char('Description')
    skip = fields.Boolean('Skip?')
    model = fields.Char('Odoo Model')
    sync_type = fields.Selection(string="Sync Type", selection=[('scheduled', 'Scheduled'), ('manual', 'Manual')])
    user_id = fields.Many2one('res.users', default=lambda self: self.env.user)


