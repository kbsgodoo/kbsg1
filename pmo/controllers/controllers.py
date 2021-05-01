# -*- coding: utf-8 -*-
from odoo import http


class Pmo(http.Controller):
      @http.route('/pmo/pmo/', auth='public')
      def index(self, **kw):
          return "Hello, world"

      @http.route('/pmo/pmo/objects/', auth='public')
      def list(self, **kw):
          return http.request.render('pmo.listing', {
              'root': '/pmo/pmo',
              'objects': http.request.env['pmo.pmo'].search([]),
          })

      @http.route('/pmo/pmo/objects/<model("pmo.pmo"):obj>/', auth='public')
      def object(self, obj, **kw):
          return http.request.render('pmo.object', {
              'object': obj
          })
