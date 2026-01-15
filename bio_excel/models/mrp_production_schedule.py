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
    def action_export_product_demand(self, ids=None):
        """Export Product Demand data to Excel

        Args:
            ids: List of production schedule IDs to export
        """
        if not xlsxwriter:
            raise UserError(_('Please install xlsxwriter python library to use this feature.\n'
                            'Command: pip install xlsxwriter'))

        # Get production schedules - use provided IDs or all if empty
        if ids:
            production_schedule_ids = self.browse(ids)
        else:
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

        # Collect all unique dates (periods) from all selected production schedules
        all_dates = set()
        for prod_schedule in production_schedule_ids:
            for forecast in prod_schedule.forecast_ids:
                if forecast.date:
                    all_dates.add(forecast.date)

        # Sort dates
        sorted_dates = sorted(list(all_dates))

        if not sorted_dates:
            raise UserError(_('No forecast dates found in selected production schedules.'))

        # Write header row
        col = 0
        worksheet.write(0, col, 'Product Internal Reference', header_format)
        col += 1
        worksheet.write(0, col, 'Product Name', header_format)
        col += 1

        # Write period date headers
        for date in sorted_dates:
            worksheet.write(0, col, date, date_format)
            col += 1

        # Add Total column
        worksheet.write(0, col, 'Total', header_format)
        total_col = col

        # Set column widths
        worksheet.set_column(0, 0, 20)  # Product Internal Reference
        worksheet.set_column(1, 1, 30)  # Product Name
        worksheet.set_column(2, total_col, 12)  # Date columns and Total

        # Write data rows - one row per production schedule
        row = 1
        for prod_schedule in production_schedule_ids:
            product = prod_schedule.product_id

            # Create a dict of date -> forecast_qty for quick lookup
            forecast_by_date = {}
            for forecast in prod_schedule.forecast_ids:
                if forecast.date:
                    forecast_by_date[forecast.date] = forecast.forecast_qty or 0.0

            # Write product info
            col = 0
            worksheet.write(row, col, product.default_code or '', data_format)
            col += 1
            worksheet.write(row, col, product.name or '', data_format)
            col += 1

            # Write forecast quantities for each period
            total_forecast = 0.0
            for date in sorted_dates:
                forecast_qty = forecast_by_date.get(date, 0.0)
                worksheet.write(row, col, forecast_qty, number_format)
                total_forecast += forecast_qty
                col += 1

            # Write total
            worksheet.write(row, total_col, total_forecast, number_format)
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
