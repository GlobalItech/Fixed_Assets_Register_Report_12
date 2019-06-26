from odoo import models, fields, api
import base64
import xlwt
from datetime import datetime,date
# from StringIO import StringIO
from io import BytesIO
import string
import math

class AccountAssetAsset(models.Model):
    _inherit = 'account.asset.asset'
    #x_studio_field_rjzDg =fields.Float('Cost Value') #that field is made from odoo studio
    
    cost_value =fields.Float('Cost Value')
    

class WizardAssetAssetHistory(models.TransientModel):
    _name='wizard.asset.asset.history'
    name = fields.Char('Report Name')
    report_file = fields.Binary('File')
    xlsx_date_from = fields.Date('Date From')
    xlsx_date_to = fields.Date('Date To')
    assest_categ_ids = fields.Many2many('account.asset.category')
    visible = fields.Boolean(default=True) #To hide the button and payslip_batch field after excel is created.
    
    @api.multi
    def get_atr(self,asset_lines):
        for asscet_line in asset_lines:
            depreciation_records=asscet_line.depreciation_line_ids.search([('asset_id.id','=',asscet_line.id),('depreciation_date','>=',self.xlsx_date_from),('depreciation_date','<=',self.xlsx_date_to)])
            quarterSum ={}
            qtr_1st = 0.0
            qtr_2nd = 0.0
            qtr_3rd = 0.0
            qtr_4th = 0.0
            for depreciation_record in depreciation_records:
                date=depreciation_record.depreciation_date
                date = datetime.strptime(date,"%Y-%m-%d").month
                quarter=q=math.ceil(date/3.)
                
                if quarter == 1: 
                    qtr_1st += depreciation_record.amount
                    quarterSum['qtr_1st'] = qtr_1st
                if quarter == 2: 
                    qtr_2nd += depreciation_record.amount
                    quarterSum['qtr_2nd'] = qtr_2nd
                if quarter == 3: 
                    qtr_3rd += depreciation_record.amount
                    quarterSum['qtr_3rd'] = qtr_3rd
                if quarter == 4: 
                    qtr_4th += depreciation_record.amount
                    quarterSum['qtr_4th'] = qtr_4th
                    
        return quarterSum
    @api.multi
    def export_asset_xls(self,asset_catg_id):
        workbook= xlwt.Workbook()
        for asset_categ_id in self.assest_categ_ids:
            asset_catg_ids = self.env['account.asset.asset'].search([('category_id','=',asset_categ_id.id),
                                                                    ('date','>=',self.xlsx_date_from),
                                                                    ('date','<=',self.xlsx_date_to)])
            fl = BytesIO()
            style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
            style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
            
           
            worksheet = workbook.add_sheet(asset_categ_id.name)
            font = xlwt.Font()
            
            font.bold = True
            for_date = xlwt.easyxf("font: name  Verdana, color black, height 200;  align: horiz left,vertical center; borders: top thin, bottom thin, left thin, right thin ; pattern: pattern solid, fore_color %s;" % '100')
            for_work_location = xlwt.easyxf("font: name  Verdana, color black, height 200;  align: horiz left,vertical center; borders: top thin, bottom thin, left thin, right medium ; pattern: pattern solid, fore_color %s;" % '100')
            for_bottom_left = xlwt.easyxf("font: name  Verdana, color black, height 200;  align: horiz left,vertical center; borders: top thin, bottom medium, left thin, right medium ; pattern: pattern solid, fore_color %s;" % '100')
    
            for_center = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom thin, left thin, right thin; pattern: pattern solid, fore_color %s;" % '100')
            for_center_right = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom thin, left thin, right medium; pattern: pattern solid, fore_color %s;" % '100')
            for_center_left = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom thin, left medium, right thin; pattern: pattern solid, fore_color %s;" % '100')

            
            for_string = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz left,vertical center; borders: top thin, bottom thin, left thin, right thin; pattern: pattern solid, fore_color %s;" % '100')
            for_last_col = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom thin, left thin, right medium; pattern: pattern solid, fore_color %s;" % '100')
            for_last_row = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom medium, left thin, right thin; pattern: pattern solid, fore_color %s;" % '100')
            for_last_row_col = xlwt.easyxf("font: name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top thin, bottom medium, left thin, right medium; pattern: pattern solid, fore_color %s;" % '100')
            for_center_heading = xlwt.easyxf("font:bold 1, name  Verdana, color black,  height 200; align: horiz center,vertical center; borders: top medium, bottom medium, left medium, right medium ")
            
            for_normal_border = xlwt.easyxf("font:bold 1, name Verdana, color black, height 200; align: horiz center, vertical center; borders: top medium, bottom medium, left medium, right medium; pattern: pattern solid, fore_color %s;" % '100')
            for_no_border = xlwt.easyxf("font: name Verdana, color black, height 200; align: horiz center, vertical center; borders: top thin, bottom thin, left thin, right thin; pattern: pattern solid, fore_color %s;" % '100')
            
            alignment = xlwt.Alignment()  # Create Alignment
            alignment.horz = xlwt.Alignment.HORZ_RIGHT
            style = xlwt.easyxf('align: wrap yes; borders: top thin, bottom thin, left thin, right thin;')
            style.num_format_str = '#,##0.00'
            
            style_net_sal = xlwt.easyxf('font:bold 1; align: wrap yes; borders: top medium, bottom medium, left medium, right medium;')
            style_net_sal.num_format_str = '#,##0.00'
            
            for limit in range(1,65536):
                worksheet.row(limit).height = 400
    
            worksheet.row(0).height = 300
            worksheet.col(0).width = 2000
            worksheet.col(1).width = 6000
            worksheet.col(2).width = 3000
            worksheet.col(3).width = 6000
            worksheet.col(4).width = 7000
            worksheet.col(5).width = 4000
            worksheet.col(6).width = 5000
            worksheet.col(7).width = 4000
            worksheet.col(8).width = 4500
            worksheet.col(9).width = 4000
            worksheet.col(10).width = 4000
            worksheet.col(11).width = 4000
            worksheet.col(12).width = 4000
            worksheet.col(13).width = 4000
            worksheet.col(14).width = 4000
            worksheet.col(15).width = 4000
            worksheet.col(16).width = 4000
            worksheet.col(17).width = 4000
            
            borders = xlwt.Borders()
            borders.bottom = xlwt.Borders.MEDIUM
            border_style = xlwt.XFStyle()  # Create Style
            border_style.borders = borders
            inv_name_row = 6
            
            worksheet.write(0, 0, 'Your Company Name', style0)
            worksheet.write(1, 0, 'Fixed Assets Register',style0)
            worksheet.write(2, 0, asset_categ_id.name, style0)
            
            worksheet.write(3, 0, 'For the Period ('  + datetime.strptime(str(self.xlsx_date_from), '%Y-%m-%d').strftime('%m/%d/%y') + '--' + datetime.strptime(str(self.xlsx_date_to), '%Y-%m-%d').strftime('%m/%d/%y') + ')',for_date)
#             worksheet.write(inv_name_row, 0, 'Purchase',for_center)
            worksheet.write_merge(inv_name_row, 7, 0,0 ,'Purchase',for_center)
            worksheet.write_merge(inv_name_row, 7, 1, 1, 'Supplier',for_center)
            worksheet.write_merge(inv_name_row, 7, 2, 2, 'Description',for_center)
            worksheet.write_merge(inv_name_row, 7, 3, 3, 'Location',for_center_left)
            worksheet.write_merge(inv_name_row, 7,4,4, datetime.strptime(str(self.xlsx_date_from), '%Y-%m-%d').strftime('%m/%d/%y') ,for_date)
            worksheet.write(inv_name_row, 5, 'C',for_center)
            worksheet.write(inv_name_row, 6, 'O',for_center)
            worksheet.write(inv_name_row, 7, 'S',for_center)
            worksheet.write(inv_name_row, 8, 'T',for_center)
            worksheet.write(inv_name_row, 9, ' ',for_center)
            worksheet.write(inv_name_row, 10, 'Rate',for_center)
            worksheet.write(inv_name_row, 11, 'Accummulated Depreciation',for_center)
            worksheet.write_merge(inv_name_row, inv_name_row, 12, 15, "Depreciation: " ,for_center)
            worksheet.write(inv_name_row, 16, 'Number Of Month',for_center)
            worksheet.write(inv_name_row, 17, 'Total',for_center)
            worksheet.write(inv_name_row, 18, 'Accumulated',for_center)
            worksheet.write(inv_name_row, 19, 'Accumulated Depreciation',for_center)
            worksheet.write(inv_name_row, 20, 'WDV',for_center)
            
            inv_name_row2 = 7
#             worksheet.write(inv_name_row2, 4, datetime.strptime(self.xlsx_date_from, '%Y-%m-%d').strftime('%m/%d/%y') ,for_date)
            worksheet.write(inv_name_row2, 5, 'Addition',for_string)
            worksheet.write(inv_name_row2, 6, 'Deletion',for_string)
            worksheet.write(inv_name_row2, 7, 'Revaluation',for_string)
            worksheet.write(inv_name_row2, 8, 'impairment',for_string)
            worksheet.write(inv_name_row2, 9,  datetime.strptime(str(self.xlsx_date_to), '%Y-%m-%d').strftime('%m/%d/%y'),for_date)
            worksheet.write(inv_name_row2, 10, '%',for_center_right)
            worksheet.write(inv_name_row2, 11,datetime.strptime(str(self.xlsx_date_from), '%Y-%m-%d').strftime('%m/%d/%y'),for_date)
            worksheet.write(inv_name_row2, 12, '1st QTR',for_center)
            worksheet.write(inv_name_row2, 13, '2nd QTR',for_center)
            worksheet.write(inv_name_row2, 14, '3rd QTR',for_center)
            worksheet.write(inv_name_row2, 15, '4th QTR',for_center)
            worksheet.write(inv_name_row2, 16, 'Used During Year',for_center)
            worksheet.write(inv_name_row2, 17, 'Total Of Filtered dates',for_center)
            worksheet.write(inv_name_row2, 18, 'Depreciation Adj',for_center)
            worksheet.write(inv_name_row2, 19, datetime.strptime(str(self.xlsx_date_to), '%Y-%m-%d').strftime('%m/%d/%y'),for_date)
            worksheet.write(inv_name_row2, 20, datetime.strptime(str(self.xlsx_date_to), '%Y-%m-%d').strftime('%m/%d/%y'),for_date)
    
    
            #adding information on sheet
            inv_name_row3 = 9
            for record in asset_catg_ids:
                if record:
                    purchase = record.x_studio_field_GPSJG or ' '
                    supplier = record.partner_id.name or ''
                    description = record.name
                    location = record.x_studio_field_EZFjg
                    asset_cost = record.cost_value #cost value
                    rate = (record.method_progress_factor * 100) or ''
                    accumulated_depr_from = record.x_studio_field_6a8eF or 0.0
                    
                    qtr_calculation =self.get_atr(record)
                    
                    qtr_1st = qtr_calculation.get('qtr_1st') or 0.0
                    qtr_2nd = qtr_calculation.get('qtr_2nd') or 0.0
                    qtr_3rd = qtr_calculation.get('qtr_3rd') or 0.0
                    qtr_4th = qtr_calculation.get('qtr_4th') or 0.0
                    year_month = 12
                    
                    total_of_yr = qtr_1st + qtr_2nd + qtr_3rd + qtr_4th
                    accumulated_depr_to = asset_cost + total_of_yr 
                    WDV =  accumulated_depr_to - asset_cost
                    
                    worksheet.write(inv_name_row3, 0, purchase,for_center)
                    worksheet.write(inv_name_row3, 1, supplier,for_string)
                    worksheet.write(inv_name_row3, 2, description,for_string)
                    worksheet.write(inv_name_row3, 3, location,for_center)
                    worksheet.write(inv_name_row3, 4, asset_cost,for_center)
                    worksheet.write(inv_name_row3, 5, ' ',for_center)
                    worksheet.write(inv_name_row3, 6, ' ',for_center)
                    worksheet.write(inv_name_row3, 7, ' ',for_center)
                    worksheet.write(inv_name_row3, 8, ' ',for_center)
                    worksheet.write(inv_name_row3, 9, ' ',for_center)
                    worksheet.write(inv_name_row3, 10, rate,for_center_left,)
                    worksheet.write(inv_name_row3, 11, accumulated_depr_from,style)
                    #This Quarters is set if finacial year is july to Jun (qtr_3rd,,qtr_1st,qtr_2nd)
                    #if finacial year is Jun to Dec then just below field change according to (qtr_1st,qtr_2nd,qtr_3rd,qtr_th) 
                    worksheet.write(inv_name_row3, 12, qtr_3rd,style)
                    worksheet.write(inv_name_row3, 13, qtr_4th,style)
                    worksheet.write(inv_name_row3, 14, qtr_1st,style)
                    worksheet.write(inv_name_row3, 15, qtr_2nd,style)
                    worksheet.write(inv_name_row3, 16, year_month,for_center)
                    worksheet.write(inv_name_row3, 17, total_of_yr,style)
                    worksheet.write(inv_name_row3, 18, ' ',style)
                    worksheet.write(inv_name_row3, 19, accumulated_depr_to,style)
                    worksheet.write(inv_name_row3, 20, WDV,style)
                    
                    inv_name_row3 +=1 
            
        workbook.save(fl)
        fl.seek(0)
        self.write({
                    'report_file': base64.encodestring(fl.getvalue()),
                    'name': 'Assets.xls'})
        self.visible = False
        return {
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'wizard.asset.asset.history',
                'target': 'new',
                'res_id': self.id,
        }
        

        
