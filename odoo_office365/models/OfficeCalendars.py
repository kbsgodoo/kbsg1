# -*- coding: utf-8 -*-

import logging
import re

from odoo import _, fields, models

_logger = logging.getLogger(__name__)
_image_dataurl = re.compile(r'(data:image/[a-z]+?);base64,([a-z0-9+/]{3,}=*)([\'"])', re.I)


class OfficeCalendars(models.Model):
    _name = "office.calendars"
    _description = "office365/Calendars"

    calendar_id = fields.Char(string="Office Calendar ID", required=False, )
    name = fields.Char(string="Calendar Name", required=False, )
    res_user = fields.Many2one(comodel_name="res.users", string="User", required=False, )
