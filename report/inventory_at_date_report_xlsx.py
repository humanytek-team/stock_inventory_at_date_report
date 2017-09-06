# -*- coding: utf-8 -*-
###############################################################################
#
#    Odoo, Open Source Management Solution
#    Copyright (C) 2017 Humanytek (<www.humanytek.com>).
#    Manuel Márquez <manuel@humanytek.com>
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools.translate import _


class InventoryAtDateReportXlsx(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, stock_history):
        import logging
        _logger = logging.getLogger(__name__)

        _logger.debug('DEBUG GENERATE XLSX REPORT SELF %s', self)
        _logger.debug('DEBUG GENERATE XLSX REPORT WORKBOOK %s', workbook)
        _logger.debug('DEBUG GENERATE XLSX REPORT DATA %s', data)
        _logger.debug('DEBUG GENERATE XLSX REPORT stock_history %s', stock_history)
        _logger.debug('DEBUG GENERATE XLSX REPORT len stock_history %s', len(stock_history))

        report_name = _('Inventory at Date')
        sheet = workbook.add_worksheet(report_name)
        bold = workbook.add_format({'bold': True})

        # Header
        sheet.write(0, 0, _('Company'), bold)
        sheet.write(0, 1, _('Product'), bold)
        sheet.write(0, 2, _('Category'), bold)
        sheet.write(0, 3, _('Quantity'), bold)
        sheet.write(0, 4, _('Price'), bold)
        sheet.write(0, 5, _('Inventory Value'), bold)

        data = list()
        product_ids = list()

        for line in stock_history:

            if line.product_id.id not in product_ids:
                data_product = dict()
                data_product[str(line.product_id.id)] = list()
                data_product[str(line.product_id.id)].append({
                    'company_id': line.company_id.id,
                    'company': line.company_id.name,
                    'product': line.product_id.name,
                    'category': line.product_id.categ_id.name,
                    'qty': line.quantity,
                    'price': line.product_id.lst_price,
                    'inventory_value': line.inventory_value,
                })
                data.append(data_product)

            else:
                product_id = str(line.product_id.id)
                data_product = False
                data_product = (item for item in data
                    if product_id in item).next()

                if data_product:
                    companies_ids = [
                        data_by_company['company_id']
                        for data_by_company in data_product[product_id]]

                    if line.company_id.id in companies_ids:
                        for data_by_company in data_product[product_id]:
                            if data_by_company['company_id'] == \
                                line.company_id.id:

                                data_by_company['qty'] += line.quantity
                                data_by_company['inventory_value'] += \
                                    line.inventory_value

                    else:
                        data_product[product_id].append({
                            'company_id': line.company_id.id,
                            'company': line.company_id.name,
                            'product': line.product_id.name,
                            'category': line.product_id.categ_id.name,
                            'qty': line.quantity,
                            'price': line.product_id.lst_price,
                            'inventory_value': line.inventory_value,
                        })

            product_ids.append(line.product_id.id)

        row = 1
        for data_product in data:
            _logger.debug('DEBUG ITER OVER DATA %s', data_product)
            data_by_company = data_product[data_product.keys()[0]]
            for data_stock_company in data_by_company:
                _logger.debug('DEBUG ITER OVER DATA %s', data_stock_company)
                sheet.write(row, 0, data_stock_company['company'])
                sheet.write(row, 1, data_stock_company['product'])
                sheet.write(row, 2, data_stock_company['category'])
                sheet.write(row, 3, data_stock_company['qty'])
                sheet.write(row, 4, data_stock_company['price'])
                sheet.write(row, 5, data_stock_company['inventory_value'])
                row += 1

InventoryAtDateReportXlsx(
    'report.inventory.at.date.report.xlsx', 'stock.history')
