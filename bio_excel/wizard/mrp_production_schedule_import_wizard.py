from odoo import fields, models, _


class MrpProductionSheduleImportWizard(models.TransientModel):
    _name = 'bio.mrp.production.schedule.import.wizard'
    _description = 'MRP Production Shedule Import Wizard'

    manufacturing_period = fields.Selection([
        ('month', 'Monthly'),
        ('week', 'Weekly')], string="Manufacturing Period",
        default='month', required=True,
        help="Default value for the time ranges in Master Production Schedule report.")
    excel_file = fields.Binary(string='Excel File', required=True)
    filename = fields.Char(string='Filename')
    line_ids = fields.One2many(
        comodel_name='bio.mrp.production.schedule.lines.import.wizard',
        inverse_name='bio_mrp_production_schedule_wizard_id', string='Lines' )


    def action_open_wizard(self):
        return {
            'name': _('Open MRP Production Shedule Import Wizard'),
            'type': 'ir.actions.act_window',
            'view_mode': 'form',
            'res_model': 'bio.mrp.production.schedule.import.wizard',
            'target': 'new',
            'context': {},
        }


class MrpProductionSheduleLinessImportWizard(models.TransientModel):
    _name = 'bio.mrp.production.schedule.lines.import.wizard'
    _description = 'MRP Production Shedule Lines Import Wizard'

    bio_mrp_production_schedule_wizard_id = fields.Many2one(
        comodel_name='bio.mrp.production.schedule.import.wizard',
        string='MRP Production Shedule Import Wizard',)
    name = fields.Char(string='name')

