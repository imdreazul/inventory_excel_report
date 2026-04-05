# -*- coding: utf-8 -*-

import io
import base64
import logging
import pytz
import xlsxwriter
from datetime import datetime, time

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError

_logger = logging.getLogger(__name__)


class InventoryReportWizard(models.TransientModel):
    _name = 'inventory.report.wizard'
    _description = 'Inventory Excel Report Wizard'

    # ─── Wizard Fields ────────────────────────────────────────────────────────

    date_from = fields.Date(
        string='Date From',
        required=True,
        default=fields.Date.context_today,
    )
    date_to = fields.Date(
        string='Date To',
        required=True,
        default=fields.Date.context_today,
    )
    location_id = fields.Many2one(
        comodel_name='stock.location',
        string='Location Name',
        required=True,
        domain=[('usage', '=', 'internal')],
        help='Select the store/warehouse location to report on.',
    )
    company_id = fields.Many2one(
        comodel_name='res.company',
        string='Company',
        readonly=True,
        default=lambda self: self.env.company,
    )
    product_category_ids = fields.Many2many(
        comodel_name='product.category',
        relation='inv_report_wiz_category_rel',
        column1='wizard_id',
        column2='category_id',
        string='Product Category',
        help='Leave empty to include all categories.',
    )
    product_ids = fields.Many2many(
        comodel_name='product.product',
        relation='inv_report_wiz_product_rel',
        column1='wizard_id',
        column2='product_id',
        string='Product Name',
        domain=[('type', 'in', ['product', 'consu'])],
        help='Leave empty to include all products.',
    )

    # ─── Constraints ──────────────────────────────────────────────────────────

    @api.constrains('date_from', 'date_to')
    def _check_dates(self):
        for rec in self:
            if rec.date_from and rec.date_to and rec.date_from > rec.date_to:
                raise ValidationError(_('Date From must be earlier than or equal to Date To.'))

    # ─── Report Entry Point ───────────────────────────────────────────────────

    def action_generate_report(self):
        """Generate and return the Excel inventory report."""
        self.ensure_one()

        products = self._get_filtered_products()
        if not products:
            raise UserError(_(
                'No storable/consumable products found matching your filters.'
            ))

        # Convert dates to datetimes (UTC boundaries)
        date_from_dt = datetime.combine(self.date_from, time.min)   # start of day
        date_to_dt   = datetime.combine(self.date_to,   time.max)   # end of day

        # ── Opening Stock (bulk, before date_from) ────────────────────────────
        opening_map = self._compute_opening_stock(products, date_from_dt)

        # ── Stock In  (moves INTO our location during period) ─────────────────
        in_move_lines = self._get_move_lines(
            location_dest_id=self.location_id.id,
            products=products,
            date_from=date_from_dt,
            date_to=date_to_dt,
        )

        # ── Stock Out (moves FROM our location during period) ─────────────────
        out_move_lines = self._get_move_lines(
            location_id=self.location_id.id,
            products=products,
            date_from=date_from_dt,
            date_to=date_to_dt,
        )

        # ── Discover dynamic locations ────────────────────────────────────────
        in_locations  = in_move_lines.mapped('location_id')        # source locs
        out_locations = out_move_lines.mapped('location_dest_id')  # dest locs

        # ── Build per-product movement dictionaries ───────────────────────────
        in_data  = {p.id: {loc.id: 0.0 for loc in in_locations}  for p in products}
        out_data = {p.id: {loc.id: 0.0 for loc in out_locations} for p in products}

        for line in in_move_lines:
            if line.product_id.id in in_data:
                in_data[line.product_id.id][line.location_id.id] += line.quantity

        for line in out_move_lines:
            if line.product_id.id in out_data:
                out_data[line.product_id.id][line.location_dest_id.id] += line.quantity

        # ── Generate Excel ────────────────────────────────────────────────────
        excel_bytes = self._build_excel(
            products, opening_map, in_data, out_data, in_locations, out_locations
        )

        # ── Attach and return download URL ────────────────────────────────────
        filename = 'Inventory_Report_{}_{}.xlsx'.format(self.date_from, self.date_to)
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(excel_bytes),
            'res_model': self._name,
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/{}?download=true'.format(attachment.id),
            'target': 'self',
        }

    # ─── Helper: Product Filtering ────────────────────────────────────────────

    def _get_filtered_products(self):
        domain = [('type', 'in', ['product', 'consu'])]
        if self.product_category_ids:
            domain.append(('categ_id', 'in', self.product_category_ids.ids))
        if self.product_ids:
            domain.append(('id', 'in', self.product_ids.ids))
        return self.env['product.product'].search(domain, order='default_code, name')

    # ─── Helper: Move Line Query ──────────────────────────────────────────────

    def _get_move_lines(self, products, date_from, date_to,
                        location_id=None, location_dest_id=None):
        """Return done stock.move.line records matching given criteria."""
        domain = [
            ('state', '=', 'done'),
            ('date', '>=', date_from),
            ('date', '<=', date_to),
            ('product_id', 'in', products.ids),
        ]
        if location_id:
            domain.append(('location_id', '=', location_id))
        if location_dest_id:
            domain.append(('location_dest_id', '=', location_dest_id))
        return self.env['stock.move.line'].search(domain)

    # ─── Helper: Opening Stock (bulk read_group) ──────────────────────────────

    def _compute_opening_stock(self, products, date_from_dt):
        """
        Opening stock = all historical IN qty - all historical OUT qty
        at self.location_id, strictly BEFORE date_from_dt.
        Uses read_group for performance.
        """
        base_domain = [
            ('state', '=', 'done'),
            ('date', '<', date_from_dt),
            ('product_id', 'in', products.ids),
        ]

        # Inbound to our location
        in_groups = self.env['stock.move.line'].read_group(
            base_domain + [('location_dest_id', '=', self.location_id.id)],
            fields=['product_id', 'quantity:sum'],
            groupby=['product_id'],
        )
        # Outbound from our location
        out_groups = self.env['stock.move.line'].read_group(
            base_domain + [('location_id', '=', self.location_id.id)],
            fields=['product_id', 'quantity:sum'],
            groupby=['product_id'],
        )

        in_map  = {g['product_id'][0]: g['quantity'] for g in in_groups}
        out_map = {g['product_id'][0]: g['quantity'] for g in out_groups}

        return {
            p.id: (in_map.get(p.id, 0.0) - out_map.get(p.id, 0.0))
            for p in products
        }

    # ─── Helper: Excel Builder ─────────────────────────────────────────────────

    def _build_excel(self, products, opening_map, in_data, out_data,
                     in_locations, out_locations):
        """Build and return raw bytes of the .xlsx report."""

        output   = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws       = workbook.add_worksheet('Inventory Report')

        # ── Format Definitions ─────────────────────────────────────────────────
        DARK_BLUE  = '#1F3864'
        MED_BLUE   = '#2E75B6'
        RED_HDR    = '#C00000'
        TOTAL_BG   = '#D6DCE4'
        HEADER_BG  = '#F2F2F2'
        WHITE      = '#FFFFFF'
        LIGHT_BLUE = '#DEEAF1'

        def fmt(**kw):
            base = {'font_name': 'Arial', 'font_size': 10}
            base.update(kw)
            return workbook.add_format(base)

        f_title = fmt(
            bold=True, font_size=16, align='center', valign='vcenter',
            font_color=DARK_BLUE,
        )
        f_label = fmt(bold=True, font_size=10)
        f_value = fmt(font_size=10)
        f_col_hdr = fmt(
            bold=True, bg_color=DARK_BLUE, font_color=WHITE,
            border=1, align='center', valign='vcenter', text_wrap=True,
        )
        f_in_grp = fmt(
            bold=True, bg_color=MED_BLUE, font_color=WHITE,
            border=1, align='center', valign='vcenter', text_wrap=True,
        )
        f_out_grp = fmt(
            bold=True, bg_color=RED_HDR, font_color=WHITE,
            border=1, align='center', valign='vcenter', text_wrap=True,
        )
        f_in_sub = fmt(
            bold=True, bg_color='#9DC3E6', font_color=DARK_BLUE,
            border=1, align='center', valign='vcenter', text_wrap=True, font_size=9,
        )
        f_out_sub = fmt(
            bold=True, bg_color='#FF9999', font_color='#7B0000',
            border=1, align='center', valign='vcenter', text_wrap=True, font_size=9,
        )
        f_total_sub_in = fmt(
            bold=True, bg_color='#4472C4', font_color=WHITE,
            border=1, align='center', valign='vcenter', text_wrap=True,
        )
        f_total_sub_out = fmt(
            bold=True, bg_color='#FF0000', font_color=WHITE,
            border=1, align='center', valign='vcenter', text_wrap=True,
        )
        f_data_str = fmt(border=1, valign='vcenter')
        f_data_num = fmt(border=1, valign='vcenter', num_format='#,##0.00', align='right')
        f_sl       = fmt(border=1, valign='vcenter', align='center')
        f_tot_str  = fmt(bold=True, border=1, bg_color=TOTAL_BG, valign='vcenter')
        f_tot_num  = fmt(bold=True, border=1, bg_color=TOTAL_BG,
                         num_format='#,##0.00', align='right', valign='vcenter')
        f_price    = fmt(border=1, valign='vcenter', num_format='#,##0.00',
                         align='right', font_color='#375623')

        # ── Column Layout ──────────────────────────────────────────────────────
        in_loc_list  = list(in_locations)
        out_loc_list = list(out_locations)

        # Fixed columns before in:  SL | Code | Name | Category | Opening
        FIXED_LEFT = 5
        # After in locations: Total In
        # After out locations: Total Out
        # Fixed right: Cost | Sales | Closing
        FIXED_RIGHT = 3

        total_cols = (
            FIXED_LEFT
            + len(in_loc_list) + 1      # in cols + Total In
            + len(out_loc_list) + 1     # out cols + Total Out
            + FIXED_RIGHT
        )

        # Column index helpers
        # COL_SL       = 0
        COL_SL       = 1
        # COL_CODE     = 1
        COL_NAME     = 2
        COL_CATEG    = 3
        COL_OPENING  = 4
        COL_IN_START = 5
        COL_IN_TOTAL = COL_IN_START + len(in_loc_list)
        COL_OUT_START = COL_IN_TOTAL + 1
        COL_OUT_TOTAL = COL_OUT_START + len(out_loc_list)
        COL_COST     = COL_OUT_TOTAL + 1
        COL_SALES    = COL_COST + 1
        COL_CLOSING  = COL_SALES + 1

        row = 0

        # ── Title Row ──────────────────────────────────────────────────────────
        ws.set_row(row, 32)
        ws.merge_range(row, 0, row, total_cols - 1, 'Inventory Report', f_title)
        row += 1

        # ── Info Section ───────────────────────────────────────────────────────
        ws.set_row(row, 18)

        def write_pair(r, c, label, value):
            ws.write(r, c,     label, f_label)
            ws.write(r, c + 1, value, f_value)

        # write_pair(row, 0, 'Date From :', str(self.date_from))
        # write_pair(row, 3, 'Date To :', str(self.date_to))
        # user_tz = pytz.timezone(self.env.user.tz or 'UTC')
        # print_dt = datetime.now(pytz.utc).astimezone(user_tz).strftime('%Y-%m-%d %H:%M:%S')
        # write_pair(row, 6, 'Print Date & Time :', print_dt)
        # row += 1
        #
        # ws.set_row(row, 16)
        # write_pair(row, 0, 'Location :', self.location_id.complete_name or self.location_id.name)
        # write_pair(row, 3, 'Company :', self.company_id.name)
        # row += 1

        write_pair(row, 1, 'Date From :', str(self.date_from))
        write_pair(row, 4, 'Date To :', str(self.date_to))
        user_tz = pytz.timezone(self.env.user.tz or 'UTC')
        print_dt = datetime.now(pytz.utc).astimezone(user_tz).strftime('%Y-%m-%d %H:%M:%S')
        write_pair(row, 7, 'Print Date & Time :', print_dt)
        row += 1

        ws.set_row(row, 16)
        write_pair(row, 1, 'Location :', self.location_id.complete_name or self.location_id.name)
        write_pair(row, 4, 'Company :', self.company_id.name)
        row += 1

        if self.product_category_ids:
            ws.set_row(row, 16)
            write_pair(row, 1, 'Product Categories :',
                       ', '.join(self.product_category_ids.mapped('complete_name')))
            row += 1

        if self.product_ids:
            ws.set_row(row, 16)
            write_pair(row, 1, 'Products :',
                       ', '.join(self.product_ids.mapped('display_name')))
            row += 1

        row += 1  # blank separator

        # ── Table Header: Row 1 (group labels) ────────────────────────────────
        ws.set_row(row, 30)
        hdr1 = row

        # Span fixed-left headers across 2 rows (merged later with row+1)
        for col, label in [
            (COL_SL,      'SL\nNo'),
            # (COL_CODE,    'Product Code\n(Internal Ref)'),
            # (COL_NAME,    'Product Name\n(with Variant)'),
            (COL_NAME,    'Product Details'),
            (COL_CATEG,   'Product\nCategory'),
            (COL_OPENING, 'Opening\nStock'),
        ]:
            ws.merge_range(hdr1, col, hdr1 + 1, col, label, f_col_hdr)

        # Stock In group header
        if in_loc_list:
            ws.merge_range(hdr1, COL_IN_START, hdr1, COL_IN_TOTAL, 'STOCK IN', f_in_grp)
        else:
            ws.merge_range(hdr1, COL_IN_START, hdr1 + 1, COL_IN_TOTAL, 'STOCK IN\n(Total)', f_in_grp)

        # Stock Out group header
        if out_loc_list:
            ws.merge_range(hdr1, COL_OUT_START, hdr1, COL_OUT_TOTAL, 'STOCK OUT', f_out_grp)
        else:
            ws.merge_range(hdr1, COL_OUT_START, hdr1 + 1, COL_OUT_TOTAL, 'STOCK OUT\n(Total)', f_out_grp)

        # Fixed-right headers spanning 2 rows
        for col, label in [
            (COL_COST,    'Cost\nPrice'),
            (COL_SALES,   'Sales\nPrice'),
            (COL_CLOSING, 'Closing\nStock'),
        ]:
            ws.merge_range(hdr1, col, hdr1 + 1, col, label, f_col_hdr)

        row += 1  # ── Header Row 2 (sub-headers) ─────────────────────────────
        ws.set_row(row, 45)

        # Stock In sub-headers (source locations)
        for i, loc in enumerate(in_loc_list):
            ws.write(row, COL_IN_START + i,
                     loc.complete_name or loc.name, f_in_sub)
        ws.write(row, COL_IN_TOTAL, 'Total\nIn', f_total_sub_in)

        # Stock Out sub-headers (destination locations)
        for i, loc in enumerate(out_loc_list):
            ws.write(row, COL_OUT_START + i,
                     loc.complete_name or loc.name, f_out_sub)
        ws.write(row, COL_OUT_TOTAL, 'Total\nOut', f_total_sub_out)

        row += 1
        data_start_row = row

        # ── Data Rows ──────────────────────────────────────────────────────────
        sl_no = 1
        grand = {
            'opening': 0.0,
            'in_total': 0.0,
            'out_total': 0.0,
            'closing': 0.0,
        }
        in_loc_totals  = [0.0] * len(in_loc_list)
        out_loc_totals = [0.0] * len(out_loc_list)

        for product in products:
            ws.set_row(row, 18)

            opening   = opening_map.get(product.id, 0.0)
            p_in_data = in_data.get(product.id, {})
            p_out_data = out_data.get(product.id, {})

            total_in  = sum(p_in_data.get(loc.id, 0.0) for loc in in_loc_list)
            total_out = sum(p_out_data.get(loc.id, 0.0) for loc in out_loc_list)
            closing   = opening + total_in - total_out

            # Fixed-left cells
            ws.write(row, COL_SL,      sl_no,                              f_sl)
            # ws.write(row, COL_CODE,    product.default_code or '',         f_data_str)
            ws.write(row, COL_NAME,    product.display_name or product.name, f_data_str)
            ws.write(row, COL_CATEG,   product.categ_id.complete_name or product.categ_id.name or '', f_data_str)
            ws.write(row, COL_OPENING, opening,                            f_data_num)

            # Stock In per source location
            for i, loc in enumerate(in_loc_list):
                qty = p_in_data.get(loc.id, 0.0)
                ws.write(row, COL_IN_START + i, qty, f_data_num)
                in_loc_totals[i] += qty

            ws.write(row, COL_IN_TOTAL, total_in, f_data_num)

            # Stock Out per destination location
            for i, loc in enumerate(out_loc_list):
                qty = p_out_data.get(loc.id, 0.0)
                ws.write(row, COL_OUT_START + i, qty, f_data_num)
                out_loc_totals[i] += qty

            ws.write(row, COL_OUT_TOTAL, total_out, f_data_num)

            # Cost / Sales / Closing
            ws.write(row, COL_COST,    product.standard_price, f_price)
            ws.write(row, COL_SALES,   product.lst_price,      f_price)
            ws.write(row, COL_CLOSING, closing,                f_data_num)

            # Accumulate grand totals
            grand['opening']   += opening
            grand['in_total']  += total_in
            grand['out_total'] += total_out
            grand['closing']   += closing

            sl_no += 1
            row   += 1

        # ── Grand Total Row ────────────────────────────────────────────────────
        ws.set_row(row, 20)
        ws.write(row, COL_SL,      '',           f_tot_str)
        # ws.write(row, COL_CODE,    '',           f_tot_str)
        ws.merge_range(row, COL_NAME, row, COL_CATEG, 'GRAND TOTAL', f_tot_str)
        ws.write(row, COL_OPENING, grand['opening'],   f_tot_num)

        for i in range(len(in_loc_list)):
            ws.write(row, COL_IN_START + i, in_loc_totals[i], f_tot_num)
        ws.write(row, COL_IN_TOTAL, grand['in_total'], f_tot_num)

        for i in range(len(out_loc_list)):
            ws.write(row, COL_OUT_START + i, out_loc_totals[i], f_tot_num)
        ws.write(row, COL_OUT_TOTAL, grand['out_total'], f_tot_num)

        ws.write(row, COL_COST,    '', f_tot_str)
        ws.write(row, COL_SALES,   '', f_tot_str)
        ws.write(row, COL_CLOSING, grand['closing'], f_tot_num)

        # ── Column Widths ──────────────────────────────────────────────────────
        ws.set_column(COL_SL,      COL_SL,      12)
        # ws.set_column(COL_CODE,    COL_CODE,    18)
        ws.set_column(COL_NAME,    COL_NAME,    32)
        ws.set_column(COL_CATEG,   COL_CATEG,   22)
        ws.set_column(COL_OPENING, COL_OPENING, 14)

        for i in range(len(in_loc_list)):
            ws.set_column(COL_IN_START + i, COL_IN_START + i, 22)
        ws.set_column(COL_IN_TOTAL, COL_IN_TOTAL, 12)

        for i in range(len(out_loc_list)):
            ws.set_column(COL_OUT_START + i, COL_OUT_START + i, 22)
        ws.set_column(COL_OUT_TOTAL, COL_OUT_TOTAL, 12)

        ws.set_column(COL_COST,    COL_COST,    13)
        ws.set_column(COL_SALES,   COL_SALES,   13)
        ws.set_column(COL_CLOSING, COL_CLOSING, 14)

        # ── Freeze panes below header ──────────────────────────────────────────
        ws.freeze_panes(data_start_row, COL_IN_START)

        # ── Print settings ─────────────────────────────────────────────────────
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.set_paper(9)  # A4

        workbook.close()
        return output.getvalue()
