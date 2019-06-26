{
    'name': 'Fixed Assets Register Report',
    'category': 'Assets',
    'license': "LGPL",
    'summary': "Give The assets report in excel",
    'author': 'Itech Resources',
    'website': 'http://www.itechresources.net',
    'depends': [
                'base',
                'account',
                'account_asset',
 #               'report_xlsx'
                ],
    'data': [
            'wizard/report_menu.xml',
            ],
    'installable': True,
    'auto_install': False,
    'price': 30.00,
    'currency': 'EUR',
}
