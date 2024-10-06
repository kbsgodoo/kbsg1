# -*- coding: utf-8 -*-
import base64
import time
from datetime import datetime
import werkzeug
import werkzeug.exceptions
import logging
import werkzeug.utils
from odoo.service import security, model as service_model
from odoo import SUPERUSER_ID
import requests
from odoo import http
from odoo.exceptions import ValidationError, UserError
from odoo.http import request,Response
import time
_logger = logging.getLogger(__name__)


class Office365Code(http.Controller):

    @http.route("/odoo", auth="public", type='http')
    def fetch_code(self, **kwargs):
        odoo_user = request.env['res.users'].sudo().browse(int(request.env.user.id))
        if "error" in kwargs:
            raise ValidationError(kwargs['error'])
        if 'code' in kwargs:
            code = kwargs.get('code')
            response = odoo_user.generate_token(code)
            if 'token' in response:
                odoo_user.sudo().update({
                    'token': response['token'],
                    'refresh_token': response['refresh_token'],
                    'expires_in': int(round(time.time() * 1000)),
                    'code': code,
                })
                request.env.cr.commit()
                return request.render("odoo_office365.token_redirect_success_page")
            else:
                return request.render("odoo_office365.token_redirect_fail_page")
        else:
            return request.render("odoo_office365.token_redirect_fail_page")

    @http.route("/mail_webhook", auth="public",type='http', csrf=False)
    def mail_webhook(self, **kw):
        if request.httprequest.headers['Content-Type']=='text/plain; charset=utf-8':
            return Response(kw['validationToken'], status=200, headers=[('Content-Type', 'text/plain; charset=utf-8')])
        else:
            data=request.env['office.sync'].extract_webhook_message(request.jsonrequest['value'][0])
            return Response(status=200)
