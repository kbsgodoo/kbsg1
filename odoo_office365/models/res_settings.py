# -*- coding: utf-8 -*-
from odoo import models, fields, api, exceptions, _
from odoo.http import request
import requests
import json
from odoo.exceptions import ValidationError, UserError

class Office3665SettingModel(models.TransientModel):
    _inherit = 'res.config.settings'

    interval_number = fields.Integer(string="Interval Number", required=False, config_parameter='odoo_office365.interval_number')
    interval_unit = fields.Selection([
        ('minutes', 'Minutes'),
        ('hours', 'Hours'),
        ('work_days', 'Work Days'),
        ('days', 'Days'),
        ('weeks', 'Weeks'),
        ('months', 'Months'),], string='Interval Unit', config_parameter='odoo_office365.interval_unit')
    import_schedular_settings = fields.Boolean("Import Schedular",config_parameter='odoo_office365.import_schedular_settings')
    export_schedular_settings = fields.Boolean("Export Schedular",config_parameter='odoo_office365.export_schedular_settings')

    def set_values(self):
        res = super(Office3665SettingModel, self).set_values()
        return res

    def _onchange_import_schedular(self):
        try:
            IrConfigParameter = self.env['ir.config_parameter'].sudo()
            interval_number = IrConfigParameter.get_param('odoo_office365.interval_number')
            interval_unit = IrConfigParameter.get_param('odoo_office365.interval_unit')
            scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Import Office 365 Data')])
            if not scheduler:
                scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Import Office 365 Data'),
                                                        ('active', '=', False)])
            if self.import_schedular_settings:
                    scheduler.active = self.import_schedular_settings
                    scheduler.interval_number = interval_number
                    scheduler.interval_type = interval_unit
                    # scheduler.user_id=self.env.user
            else:
                scheduler.active = self.import_schedular_settings
        except Exception as e:
            raise ValidationError(str(e))

    def _onchange_export_schedular(self):
        try:
            IrConfigParameter = self.env['ir.config_parameter'].sudo()
            interval_number = IrConfigParameter.get_param('odoo_office365.interval_number')
            interval_unit = IrConfigParameter.get_param('odoo_office365.interval_unit')
            if self.export_schedular_settings:
                scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Export Odoo Data')])
                if not scheduler:
                    scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Export Odoo Data'),
                                                            ('active', '=', False)])
                scheduler.active = self.export_schedular_settings
                scheduler.interval_number = interval_number
                scheduler.interval_type = interval_unit
                # scheduler.user_id = self.env.user
            else:
                scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Export Odoo Data')])
                if not scheduler:
                    scheduler = self.env['ir.cron'].search([('name', '=', 'Auto Export Odoo Data'),
                                                            ('active', '=', False)])
                scheduler.active = self.export_schedular_settings
        except Exception as e:
            raise ValidationError(str(e))

    def start_schedulars(self):
        self._onchange_import_schedular()
        self._onchange_export_schedular()



