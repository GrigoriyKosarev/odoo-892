import base64
import io
from odoo import fields, models, _, api
from odoo.exceptions import UserError


class MrpProductionSheduleImportWizard(models.TransientModel):
    _name = 'bio.mrp.production.schedule.import.wizard'
    _description = 'MRP Production Shedule Import Wizard'

    @api.model
    def _default_warehouse_id(self):
        return self.env['stock.warehouse'].search([('company_id', '=', self.env.company.id)], limit=1)

    manufacturing_period = fields.Selection([
        ('month', 'Monthly'),
        ('week', 'Weekly')], string="Manufacturing Period",
        default='month', required=True,
        help="Default value for the time ranges in Master Production Schedule report.")
    warehouse_id = fields.Many2one('stock.warehouse', 'Production Warehouse',
        required=True, default=lambda self: self._default_warehouse_id())
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

    def action_import(self):
        """Import data from Excel file"""
        self.ensure_one()

        if not self.excel_file:
            raise UserError(_('Please upload an Excel file.'))

        # Decode the file
        try:
            file_data = base64.b64decode(self.excel_file)
            file_obj = io.BytesIO(file_data)
        except Exception as e:
            raise UserError(_('Error reading file: %s') % str(e))

        # TODO: Add Excel processing logic here
        # For now, just show success message

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('Excel file uploaded successfully. Import logic will be implemented.'),
                'type': 'success',
                'sticky': False,
            }
        }


class MrpProductionSheduleLinessImportWizard(models.TransientModel):
    _name = 'bio.mrp.production.schedule.lines.import.wizard'
    _description = 'MRP Production Shedule Lines Import Wizard'

    bio_mrp_production_schedule_wizard_id = fields.Many2one(
        comodel_name='bio.mrp.production.schedule.import.wizard',
        string='MRP Production Shedule Import Wizard',)
    name = fields.Char(string='name')

