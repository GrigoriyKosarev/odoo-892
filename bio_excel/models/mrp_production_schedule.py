import base64
from io import BytesIO
from odoo import models, fields, api, _
from odoo.exceptions import UserError
from odoo.tools.misc import xlsxwriter
# import xlsxwriter


class MrpProductionSchedule(models.Model):
    _inherit = 'mrp.production.schedule'

    excel_file = fields.Binary('Excel File', readonly=True)
    excel_filename = fields.Char('Filename', readonly=True)

    @api.model
    def action_export_product_demand(self):
        """Export Product Demand data to Excel"""
        if not xlsxwriter:
            raise UserError(_('Please install xlsxwriter python library to use this feature.\n'
                            'Command: pip install xlsxwriter'))

        # Get production schedules from context or search all
        production_schedule_ids = self.search([])

        if not production_schedule_ids:
            raise UserError(_('No production schedules found to export.'))

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Product Demand')

        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        field_label_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })

        data_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        number_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })

        date_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'dd.mm.yyyy'
        })

        # Write header row
        headers = ['Product Code', 'Product Name', 'BOM', 'Warehouse', 'Date', 'Forecast Qty', 'Indirect Demand Qty']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # Set column widths
        worksheet.set_column(0, 0, 15)  # Product Code
        worksheet.set_column(1, 1, 30)  # Product Name
        worksheet.set_column(2, 2, 30)  # BOM
        worksheet.set_column(3, 3, 20)  # Warehouse
        worksheet.set_column(4, 4, 12)  # Date
        worksheet.set_column(5, 6, 15)  # Quantities

        # Write data rows
        row = 1
        for prod_schedule in production_schedule_ids:
            product = prod_schedule.product_id
            bom = prod_schedule.bom_id
            warehouse = prod_schedule.warehouse_id

            # Get all forecast lines for this production schedule
            forecast_lines = prod_schedule.forecast_ids.sorted('date')

            if not forecast_lines:
                # If no forecasts, write one row with product info
                worksheet.write(row, 0, product.default_code or '', data_format)
                worksheet.write(row, 1, product.name or '', data_format)
                worksheet.write(row, 2, bom.display_name or '', data_format)
                worksheet.write(row, 3, warehouse.name or '', data_format)
                worksheet.write(row, 4, '', data_format)
                worksheet.write(row, 5, 0.0, number_format)
                worksheet.write(row, 6, 0.0, number_format)
                row += 1
            else:
                # Write one row per forecast line
                for forecast in forecast_lines:
                    worksheet.write(row, 0, product.default_code or '', data_format)
                    worksheet.write(row, 1, product.name or '', data_format)
                    worksheet.write(row, 2, bom.display_name or '', data_format)
                    worksheet.write(row, 3, warehouse.name or '', data_format)
                    worksheet.write(row, 4, forecast.date, date_format)
                    worksheet.write(row, 5, forecast.forecast_qty or 0.0, number_format)
                    worksheet.write(row, 6, forecast.indirect_demand_qty or 0.0, number_format)
                    row += 1

        # Close workbook and get file data
        workbook.close()
        output.seek(0)
        file_data = output.read()
        output.close()

        # Encode file to base64
        excel_file = base64.b64encode(file_data)

        # Create attachment for download
        filename = 'production_schedule_export.xlsx'
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': excel_file,
            'res_model': 'mrp.production.schedule',
            'res_id': 0,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        # Return action to download file
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }
