import base64
from io import BytesIO
from odoo import models, fields, _
from odoo.exceptions import UserError
from odoo.tools.misc import xlsxwriter
# import xlsxwriter


class MrpProductionSchedule(models.Model):
    _inherit = 'mrp.production.schedule'

    excel_file = fields.Binary('Excel File', readonly=True)
    excel_filename = fields.Char('Filename', readonly=True)

    def action_export_product_demand(self):
        """Export Product Demand data to Excel"""
        if not xlsxwriter:
            raise UserError(_('Please install xlsxwriter python library to use this feature.\n'
                            'Command: pip install xlsxwriter'))

        self.ensure_one()

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

        # Get production schedules with current filters
        production_schedule_ids = self.production_schedule_ids.filtered(
            lambda ps: self.product_id and ps.product_id == self.product_id or not self.product_id
        )

        if not production_schedule_ids:
            raise UserError(_('No data available to export. Please check your filters.'))

        # Get date ranges (periods)
        date_ranges = self._get_date_range()

        # Write header row with periods
        col = 0
        worksheet.write(0, col, '', header_format)  # Empty cell for field names

        period_columns = []
        for date_range in date_ranges:
            col += 1
            date_start = date_range[0]
            date_stop = date_range[1]
            period_label = f"{date_start.strftime('%Y-%m-%d')} - {date_stop.strftime('%Y-%m-%d')}"
            worksheet.write(0, col, period_label, header_format)
            period_columns.append((date_start, date_stop))

        # Add Total column
        col += 1
        worksheet.write(0, col, 'Total', header_format)
        total_col = col

        # Set column widths
        worksheet.set_column(0, 0, 30)  # Field name column
        worksheet.set_column(1, total_col, 15)  # Data columns

        # Process each production schedule (product)
        row = 1
        for prod_schedule in production_schedule_ids:
            product = prod_schedule.product_id

            # Row 1: Product Internal Reference
            worksheet.write(row, 0, 'Product Internal Reference', field_label_format)
            col = 1
            for _ in period_columns:
                worksheet.write(row, col, product.default_code or '', data_format)
                col += 1
            worksheet.write(row, total_col, product.default_code or '', data_format)
            row += 1

            # Row 2: Product Name
            worksheet.write(row, 0, 'Product Name', field_label_format)
            col = 1
            for _ in period_columns:
                worksheet.write(row, col, product.name or '', data_format)
                col += 1
            worksheet.write(row, total_col, product.name or '', data_format)
            row += 1

            # Row 3: Indirect Demand Forecast per period
            worksheet.write(row, 0, 'Indirect Demand Forecast per period', field_label_format)
            col = 1
            total_indirect_demand = 0.0

            # Get forecast values for each period
            for date_start, date_stop in period_columns:
                # Get forecast lines for this period
                forecast_lines = prod_schedule.forecast_ids.filtered(
                    lambda f: f.date >= date_start and f.date <= date_stop
                )

                # Sum indirect demand for this period
                period_indirect_demand = sum(forecast_lines.mapped('indirect_demand_qty'))
                worksheet.write(row, col, period_indirect_demand, number_format)
                total_indirect_demand += period_indirect_demand
                col += 1

            # Write total indirect demand
            worksheet.write(row, total_col, total_indirect_demand, number_format)
            row += 1

            # Row 4: Total Indirect Demand Forecast per period
            # This appears to be a cumulative or summary row
            # Based on the requirement, this might be the same as above or calculated differently
            worksheet.write(row, 0, 'Total Indirect Demand Forecast per period', field_label_format)
            col = 1
            cumulative_total = 0.0

            for date_start, date_stop in period_columns:
                forecast_lines = prod_schedule.forecast_ids.filtered(
                    lambda f: f.date >= date_start and f.date <= date_stop
                )
                period_total = sum(forecast_lines.mapped('indirect_demand_qty'))
                cumulative_total += period_total
                worksheet.write(row, col, cumulative_total, number_format)
                col += 1

            worksheet.write(row, total_col, cumulative_total, number_format)
            row += 1

            # Add empty row between products
            row += 1

        # Close workbook and get file data
        workbook.close()
        output.seek(0)
        file_data = output.read()
        output.close()

        # Encode file to base64
        excel_file = base64.b64encode(file_data)

        # Update wizard with file
        self.write({
            'excel_file': excel_file,
            'excel_filename': 'product_demand_report.xlsx'
        })

        # Return action to download file
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=mrp.production.schedule&id={self.id}&field=excel_file&filename={self.excel_filename}&download=true',
            'target': 'self',
        }
