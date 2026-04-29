# -*- coding: utf-8 -*-
from odoo import models
from datetime import date
import io, base64
import re
from odoo.exceptions import AccessError, UserError, ValidationError

class ReportMRPExcel(models.AbstractModel):
    _name = "report.cdp_mrp_report.cdp_mrp_report_list"
    _inherit = "report.report_xlsx.abstract"
    _description = "Excel Report - MRP Report"
    
    def contains_data(self, field_name):
        return field_name if field_name else ""

    def calculate_cell_height(self, text_data):
        if text_data:
            formatted_data = text_data.strip()
            cell_height = 15 * len(formatted_data.splitlines())
        else:
            formatted_data = ""
            cell_height = 15
        return (formatted_data, cell_height)

    def generate_xlsx_report(self, workbook, data, records):
        records.ensure_one()
        # raise UserError(f'Mohon maaf tidak bisa ..{records.name}')
        sheet = workbook.add_worksheet("Coating Progress")
        
        source = records.origin
        mo_number = records.name
        product_name = records.product_id.name
        print_quantity = records.x_studio_print_quantity
        
    
        # =====================
        # PAGE SETUP
        # =====================
        sheet.set_paper(9)        # A4
        sheet.set_landscape()     # Landscape
        # sheet.set_margins(0.5, 0.5, 0.5, 0.5)
        
        # =====================
        # STYLE
        # =====================
        title = workbook.add_format({
            "bold": 1,
            "font_size": 36,
            "align": "center",
            "valign": "vcenter",
            "font_color": "#2F75B5",
        })


        # =====================
        # COLUMN WIDTH
        sheet.set_column("A:A", 5)
        sheet.set_column("B:D", 7.45)
        sheet.set_column("E:E", 8.73)
        sheet.set_column("F:H", 10.09)
        sheet.set_column("I:I", 8.73)
        sheet.set_column("J:J", 5.73)
        sheet.set_column("K:K", 9.64)
        sheet.set_column("L:M", 6.73)
        sheet.set_column("N:N", 10.18)
        sheet.set_column("O:O", 8.64)
        sheet.set_column("P:P", 8.64)
        sheet.set_column("Q:S", 17.91)
        sheet.set_column("T:V", 10.64)
        sheet.set_column("W:W", 14.09)
        sheet.set_column("X:Y", 9.82)
        sheet.set_column("Z:AB", 12.64)
        sheet.set_column("AC:AD", 9.73)
        sheet.set_column("AE:AE", 27.64)

        # 8 = row 9 
        for r in range(8, 100): 
            sheet.set_row(r, 45)
        
        # =====================
        # TITLE
        # =====================
        # TITLE ROW
        sheet.set_row(0, 70)        
        sheet.merge_range("A1:AE1", "工单涂覆进度表", title)

        # SUBTITLE ROW
        left = workbook.add_format({
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "font_size": 26,            
            "font_color": "#2F75B5",
                        
        })
        # KOLOM 2
        # A-C
        center_header = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 26,            
            "font_color": "#2F75B5",
                        
        })        
        center_header_orange = workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 26,            
            "font_color": "#2F75B5",
            "bg_color": "#FCE4D6",
                        
        })        
        sheet.set_row(1, 42)
        sheet.write("A2", "订单号：", left)        
        
        # D-I
        sheet.set_row(1, 42)
        sheet.merge_range("D2:I2", source, center_header_orange)        

        # J-L
        sheet.set_row(1, 42)
        sheet.merge_range("J2:L2", "工单编号：", left)        

        # M-Q
        sheet.set_row(1, 42)
        sheet.merge_range("M2:Q2", mo_number, center_header_orange)        

        # R-S
        sheet.set_row(1, 42)
        sheet.merge_range("R2:S2", "产品名称：", center_header)        
  

        # T-Y
        sheet.set_row(1, 42)
        sheet.merge_range("T2:Y2", product_name, center_header_orange)        

        # Z
        sheet.set_row(1, 42)
        sheet.write("Z2:Z2", "数量:", left)        

        # AA2 - AB2
        sheet.set_row(1, 42)
        sheet.merge_range("AA2:AB2", f"{print_quantity}", center_header_orange)        

        # AC2 - AD2
        sheet.set_row(1, 42)
        sheet.merge_range("AC2:AD2", "批次号：", left)        

        # AE2 - AE2
        sheet.set_row(1, 42)
        sheet.write("AE2:AE2", "", left)  
        
        # KOLOM 3
        center_header_20 = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 20,            
            "font_color": "#2F75B5",
                        
        })          
        center_header_22 = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 22,            
            "font_color": "#2F75B5",
                        
        })          
        
        # A-H
        sheet.set_row(2, 42)
        sheet.merge_range("A3:H3", "产品涂覆指导", center_header_20)        
        
        # -I-W
        sheet.set_row(2, 42)
        sheet.merge_range("I3:W3", "涂覆班组填写", center_header_22)        

        # X3-AB3
        sheet.set_row(2, 42)
        sheet.merge_range("X3:AB3", "烧结员工填写	", center_header_22)        

        # AC3 - AD3
        sheet.set_row(2, 42)
        sheet.merge_range("AC3:AD3", "料架号：", left)        

        # AE3 - AE3
        sheet.set_row(2, 42)
        sheet.write("AE3:AE3", "", left)  

        # Section ketiga header
        # A4-A6
        center_header_16_orange = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 16,            
            "font_color": "#2F75B5",
            "bg_color": "#FCE4D6",
            "text_wrap": True
        })             
        center_header_16_gray = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 16,            
            "font_color": "#2F75B5",
            "bg_color": "#D0CECE",
            "text_wrap": True
                        
        })             
        center_header_16 = workbook.add_format({
            "bold": 1,            
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 16,            
            "font_color": "#2F75B5",
            "text_wrap": True
                        
        })         
        # Kolom 4
         
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("A4:A6", "层数", center_header_16_orange)
        # B3-B6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("B4:B6", "涂液代码", center_header_16_orange)

        # C3-C6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("C4:C6", "烧结程序", center_header_16_orange)

        # D3-D6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("D4:D6", "涂层位置", center_header_16_orange)
        
        # E3-E6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("E4:E6", "涂覆面", center_header_16_orange)
        
        # F3-F6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("F4:F6", "单片面积(m2)", center_header_16_orange)

        # G3-G6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("G4:G6", "单片指导用量", center_header_16_orange)

        # H3-H6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("H4:H6", "单层涂液用量", center_header_16_orange)

        # I3-I6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("I4:I6", "涂覆日期", center_header_16)               

        # J3-J6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("J4:J6", "涂覆班次", center_header_16)        

        # K3-K6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("K4:K6", "涂覆组", center_header_16)

        # L3-L6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("L4:L6", "涂覆人数", center_header_16)

        # M3-M6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("M4:M6", "辅助人数", center_header_16)

        # N3-N6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("N4:N6", "涂覆开始时间", center_header_16)

        # O3-O6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("O4:O6", "涂覆结束时间", center_header_16)

        # P3-P6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("P4:P6", "计划烧结开始时间", center_header_16)
        
        # Q3-U3
        sheet.set_row(3, 26)
        sheet.merge_range("Q4:U4", "姓名/数量", center_header_16_gray)

        # Q4-S4
        sheet.set_row(4, 26)
        sheet.merge_range("Q5:S5", "涂覆作业员", center_header_16_gray)

        # T4-U4
        sheet.set_row(4, 26)
        sheet.merge_range("T5:U5", "辅助作业员", center_header_16_gray)
        
        # Q5 - Q5
        sheet.set_row(5, 55.5)
        sheet.write("Q6:Q6", "A", center_header_16_gray)  

        # R5 - R5
        sheet.set_row(5, 55.5)
        sheet.write("R6:R6", "B", center_header_16_gray)  

        # S5 - S5
        sheet.set_row(5, 55.5)
        sheet.write("S6:S6", "C", center_header_16_gray)  

        # T5 - T5
        sheet.set_row(5, 55.5)
        sheet.write("T6:T6", "D", center_header_16_gray)  

        # U5 - U5
        sheet.set_row(5, 55.5)
        sheet.write("U6:U6", "E", center_header_16_gray)  
        
        # V3-V6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("V4:V6", "总工时", center_header_16_gray)        

        # W3-W6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("W4:W6", "人效(m2/人/时)", center_header_16_gray)        

        # X3-X6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("X4:X6", "当班架次", center_header_16)        

        # Y3-Y6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("Y4:Y6", "烧结程序", center_header_16)        

        # Z3-Z6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("Z4:Z6", "计划烧结开始时间", center_header_16)        

        # AA3-AA6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("AA4:AA6", "实际烧结开始时间", center_header_16)        

        # AB3-AB6
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.set_row(5, 55.5)
        sheet.merge_range("AB4:AB6", "实际烧结结束时间", center_header_16)        

        # AC4-AE5
        sheet.set_row(3, 26)
        sheet.set_row(4, 26)
        sheet.merge_range("AC4:AE5", "备注", center_header_16)
        
        # AC5 - AC5
        sheet.set_row(5, 55.5)
        sheet.write("AC6:AE6", "料架号变更", center_header_16)  

        # AD5 - AD5
        sheet.set_row(5, 55.5)
        sheet.write("AD6:AD6", "数量变更", center_header_16)  

        # AE5 - AE5
        sheet.set_row(5, 55.5)
        sheet.write("AE6:AE6", "异常备注", center_header_16)  
        
        # =====================
        # TABLE HEADER (ROW 4-6 MULTI HEADER)
        # =====================
        row = 5
        # =====================
        # DATA TABLE
        # =====================
        row += 1
        lines = []
        dn_line = records.product_tmpl_id.x_studio_dn_line or []
        for data in dn_line:
            selected = {
                "code": data.x_studio_coat.name,
                "process": data.x_studio_sintering_procedure_1,
                "position": data.x_studio_coating_position,
                "face": data.x_studio_coating_layer_1,
                "area": data.x_studio_coat_area.x_studio_coat_area or '',
                "qty_per_piece": data.x_studio_single_layer_coating_amount,
                "layer_qty": (
                    data.x_studio_coat_area.x_studio_coat_area * data.x_studio_single_layer_coating_amount
                    if data.x_studio_coat_area and data.x_studio_single_layer_coating_amount
                    else ''
                ),
            }
            lines.append(selected)
            
        start_row = row
        MAX_ROW = 50
        for i in range(MAX_ROW):
            l = lines[i] if i < len(lines) else {}
            values = [
                i + 1,
                l.get("code", ""),
                l.get("process", ""),
                l.get("position", ""),
                l.get("face", ""),
                l.get("area", ""),
                l.get("qty_per_piece", ""),
                l.get("layer_qty", ""),
                l.get("date", ""),
                l.get("shift", ""),
                l.get("group", ""),
                l.get("people", ""),
                l.get("assist", ""),
                l.get("start", ""),
                l.get("end", ""),
                l.get("plan_sinter", ""),
                l.get("A", ""),
                l.get("B", ""),
                l.get("C", ""),
                l.get("D", ""),
                l.get("E", ""),
                l.get("total_hours", ""),
                l.get("efficiency", ""),
                l.get("shift_count", ""),
                l.get("sinter_process", ""),
                l.get("plan_start", ""),
                l.get("actual_start", ""),
                l.get("actual_end", ""),
                l.get("remark", ""),
                l.get("ad_val", ""),
                l.get("ae_val", ""),
            ]

            for c, v in enumerate(values):
                fmt = center_header_16_orange if c <= 7 else center_header_16
                sheet.write(start_row + i, c, v, fmt)