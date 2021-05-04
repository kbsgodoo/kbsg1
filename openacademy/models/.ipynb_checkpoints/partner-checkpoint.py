# -*- coding: utf-8 -*-
from odoo import fields, models, api

class res_partner(models.Model):
    _name = 'openacademy.alfa'
    _inherits = {'res.partner':'beta_id'}
    
    beta_id = fields.Many2one('res.partner',ondelete='set null', string="Responsible2")

    # Add a new column to the res.partner model, by default partners are not instructors
    instructor = fields.Boolean("Instructor", default=False)
    #session_ids = fields.Many2many('openacademy.session',  string="Attended Sessions", readonly=True)
