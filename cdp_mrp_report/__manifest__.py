# -*- coding: utf-8 -*-
{
    "name": "MRP Report for (Powerstone)",
    "version": "16.0.0.0.0",
    "category": "Services",
    "author": "Falinwa",
    "summary": " MRP Report in Excel Format",
    "description": """""",
    "depends": [
        'base',
        'mail',
        'sale',
        'product',
        'report_xlsx',
    ],    
    "data": [
        "report/cdp_mrp_report_list.xml",
        "report/ir_actions_report.xml",    
        "report/ir_actions_report_templates.xml",
    ],    
    "demo": [],
    "application": False,
    "installable": True,
    "auto_install": False,
    "license": "LGPL-3",
}
