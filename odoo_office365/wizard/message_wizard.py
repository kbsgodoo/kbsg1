from  odoo import fields, models,api


class CustomMessageWizard(models.TransientModel):
    _name = 'message.wizard'
    _description = "Message/Wizard"

    title =fields.Char()

    def get_default(self):
        if self.env.context.get("message",False):
            return self.env.context.get("message")
        return False

    text = fields.Text('Message', readonly=True, default=get_default)





