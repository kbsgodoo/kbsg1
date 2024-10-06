# -*- coding: utf-8 -*-
{
    'name': "Odoo Office 365 Connector",

    'summary': """
                Odoo Office365 Connector provides the opportunity to sync calendar, contacts, tasks and mails between ODOO and Office365.
            """,

    'description': """
    odoo office365 connector
    Odoo Office 365 Connector
    Odoo Microsoft Office 365 Integration
    Office 365 with Odoo integration
    office365 odoo connector
    Microsoft
    Microsoft outlook
    Microsoft Office 365
    Microsoft Office365
    Microsoft Office365 Connector
    Microsoft Office365 integration
    office365
    office
    office 365 odoo connector
    email sync 
    sync
    email
    calendar
    calender sync 
    outlook calender syc
    outlook odoo connector
    odoo outlook connector
    odoo office bridge
    office odoo bridge
    """,
    'author': "Techloyce",
    'website': "http://www.techloyce.com",
    'category': 'Connector',
    'price': 499,
    'currency': 'USD',
    'version': '17.2.0',
    'license': 'OPL-1',
    'depends': ['base', 'calendar', 'crm'],
    'images': [
        'static/description/banner.gif',
    ],
    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'views/template.xml',
        'views/office365_sync.xml',
        'views/logging.xml',
        'views/res_partner.xml',
        'views/todo_list.xml',
        'views/Webhooks.xml',
        'views/calendar_event.xml',
        'views/res_settings.xml',
        'views/office_setting.xml',
        'views/officeImportExceptions.xml',
        'views/officeExportExceptions.xml',
        'data/scheduler.xml',
        'data/general.xml',
        'data/recurring.xml',
        'wizard/message_wizard.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
        'demo/demo.xml',
    ],
    'live_test_url': 'https://www.youtube.com/watch?v=gvuUiAsC-TM',
}
