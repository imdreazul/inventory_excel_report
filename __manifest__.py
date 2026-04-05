# -*- coding: utf-8 -*-
{
    'name': 'Inventory Excel Report',
    'summary': 'Generate dynamic inventory Excel reports with stock movements',
    'description': """
        Inventory Excel Report
        ======================
        Generate professional Excel reports for inventory with:
        - Opening Stock (before date range)
        - Dynamic Stock In columns per source location
        - Dynamic Stock Out columns per destination location
        - Closing Stock calculation
        - Cost Price & Sales Price
        - Filters: Date Range, Location, Company, Product Category, Product
    """,
    'author': "Reazul",
    'website': "https://github.com/imdreazul",
    'category': 'Generic Modules/Inventory Report',
    'version': '1.2',

    'depends': ['stock', 'product'],
    'data': [
        'security/ir.model.access.csv',
        'views/inventory_report_wizard_view.xml',
    ],
    'installable': True,
    'auto_install': False,
    'license': 'LGPL-3',
}
