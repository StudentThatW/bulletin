# -*- coding: utf-8 -*-
{

    'name': "bulletin",
    'summary': "module bulletin",
    'description': "module bulletin",

    'author': "Ids",
    'website': "",

    'category': 'Sales',

    # порядковый № отображения в списке приложений
    'sequence': 2,
    'version': '0.1',

    'installable': True,
    # является приложением
    'application': True,
    'auto_install': False,

    'images': ['static/description/sign.png'],

    # any module necessary for this one to work correctly
    'depends': ['base'],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        'views/views.xml',
        'views/templates.xml',
    ],
    # only loaded in demonstration mode
    # 'demo': [
    #     'demo/demo.xml',
    # ],
}
