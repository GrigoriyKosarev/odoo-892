import base64
import io
from datetime import datetime
from odoo import fields, models, api, _
from odoo.exceptions import UserError

try:
    import openpyxl
except ImportError:
    openpyxl = None


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

    def action_upload(self):
        """Parse Excel file and create wizard lines"""
        self.ensure_one()

        if not self.excel_file:
            raise UserError(_('Please upload an Excel file.'))

        if not openpyxl:
            raise UserError(_('Python library "openpyxl" is not installed. Please install it: pip install openpyxl'))

        # Decode the file
        try:
            file_data = base64.b64decode(self.excel_file)
            file_obj = io.BytesIO(file_data)
        except Exception as e:
            raise UserError(_('Error reading file: %s') % str(e))

        # Parse Excel file
        try:
            workbook = openpyxl.load_workbook(file_obj, data_only=True)
            sheet = workbook.active
        except Exception as e:
            raise UserError(_('Error parsing Excel file: %s') % str(e))

        # Clear existing lines
        self.line_ids.unlink()

        # Parse header row (dates)
        header_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]

        # Extract dates from header (skip first 2 columns: default_code, Назва)
        date_columns = []
        for col_idx, cell_value in enumerate(header_row[2:], start=2):
            if cell_value:
                try:
                    # Try to parse date from DD.MM.YYYY format
                    if isinstance(cell_value, str):
                        forecast_date = datetime.strptime(cell_value.strip(), '%d.%m.%Y').date()
                    elif isinstance(cell_value, datetime):
                        forecast_date = cell_value.date()
                    else:
                        continue
                    date_columns.append((col_idx, forecast_date))
                except ValueError:
                    # Skip invalid dates
                    continue

        if not date_columns:
            raise UserError(_('No valid dates found in header row. Expected format: DD.MM.YYYY (e.g., 01.12.2025)'))

        # Parse data rows (starting from row 2)
        lines_to_create = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            default_code = row[0]
            product_name = row[1] if len(row) > 1 else ''

            if not default_code:
                continue  # Skip empty rows

            # Find product by default_code
            product = self.env['product.product'].search([('default_code', '=', default_code)], limit=1)

            # Create line for each date column
            for col_idx, forecast_date in date_columns:
                qty = row[col_idx] if col_idx < len(row) else 0

                # Skip if quantity is 0 or empty
                if not qty:
                    continue

                try:
                    forecast_qty = float(qty)
                except (ValueError, TypeError):
                    forecast_qty = 0.0

                if forecast_qty <= 0:
                    continue

                line_vals = {
                    'bio_mrp_production_schedule_wizard_id': self.id,
                    'default_code': str(default_code),
                    'product_name': str(product_name) if product_name else '',
                    'forecast_date': forecast_date,
                    'forecast_qty': forecast_qty,
                }

                if product:
                    line_vals.update({
                        'product_id': product.id,
                        'state': 'found',
                        'message': _('Product found'),
                    })
                else:
                    line_vals.update({
                        'state': 'not_found',
                        'message': _('Product with code "%s" not found in database') % default_code,
                    })

                lines_to_create.append(line_vals)

        # Create lines
        if lines_to_create:
            self.env['bio.mrp.production.schedule.lines.import.wizard'].create(lines_to_create)

        # Return view with lines
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'bio.mrp.production.schedule.import.wizard',
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_import(self):
        """Import validated lines to mrp.production.schedule"""
        self.ensure_one()

        lines_to_import = self.line_ids.filtered(lambda l: l.state == 'found')

        if not lines_to_import:
            raise UserError(_('No valid lines to import. Please upload and validate Excel file first.'))

        # TODO: Implement import to mrp.production.schedule
        # Group by product and create/update production schedule records

        imported_count = len(lines_to_import)

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('%d lines imported successfully. Import logic will be fully implemented soon.') % imported_count,
                'type': 'success',
                'sticky': False,
            }
        }


class MrpProductionSheduleLinessImportWizard(models.TransientModel):
    _name = 'bio.mrp.production.schedule.lines.import.wizard'
    _description = 'MRP Production Shedule Lines Import Wizard'
    _order = 'default_code, forecast_date'

    bio_mrp_production_schedule_wizard_id = fields.Many2one(
        comodel_name='bio.mrp.production.schedule.import.wizard',
        string='Wizard',
        required=True,
        ondelete='cascade')

    # Product information
    product_id = fields.Many2one(
        'product.product',
        string='Product',
        help='Product found by default_code')
    default_code = fields.Char(
        string='Product Code',
        required=True,
        help='Product code from Excel file (Column A)')
    product_name = fields.Char(
        string='Product Name',
        help='Product name from Excel file (Column B)')

    # Forecast data
    forecast_date = fields.Date(
        string='Forecast Date',
        required=True,
        help='Date from Excel header (DD.MM.YYYY format)')
    forecast_qty = fields.Float(
        string='Forecast Quantity',
        digits='Product Unit of Measure',
        required=True,
        help='Quantity from Excel cell')

    # Status
    state = fields.Selection([
        ('draft', 'Draft'),
        ('found', 'Product Found'),
        ('not_found', 'Product Not Found'),
        ('imported', 'Imported'),
        ('error', 'Error'),
    ], string='Status', default='draft', required=True)

    message = fields.Text(
        string='Message',
        help='Status message or error description')

