# -*- coding: utf-8 -*-
from odoo import http


class Modulo4(http.Controller):
     @http.route('/modulo4/modulo4/', auth='public')
     def index(self, **kw):
        return "Hello, world"

     @http.route('/modulo4/modulo4/objects/', auth='public')
     def list(self, **kw):
         return http.request.render('modulo4.listing', {
             'root': '/modulo4/modulo4',
             'objects': http.request.env['modulo4.modulo4'].search([]),
         })

     @http.route('/modulo4/modulo4/objects/<model("modulo4.modulo4"):obj>/', auth='public')
     def object(self, obj, **kw):
         return http.request.render('modulo4.object', {
             'object': obj
         })
