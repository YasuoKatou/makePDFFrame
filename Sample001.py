# -*- coding: utf-8 -*-

import pathlib

from PDFPreview import makePDFwithExcel

if __name__ == '__main__':
    excelFile = 'V01-frame_100_LibreOffice.xlsx'
    excelPath = pathlib.Path(__file__).parent / 'samples' / excelFile
    jsonPath  = str(excelPath.parent / excelPath.stem) + '.json'
    mkInfo = {
        'JsonPath' : str(jsonPath),
        'ExcelPath': str(excelPath),
        'SheetName': 'sample001',
        'PDFPath'  : '',
        'JsonOut'  : False,
        '_debug'   : False,
        'args'     : {
            'page_margins': {
                'top':  20,
                'left': 20
            },
            'row_height': {
                'width':  5,
                'height': 5
            }
        },
        'data': {
            'date1': '2021年 5月31日',
            'date2': '2099年 7月末日',
            'name1': '日本国政府 株式会社',
            'amount': '300,000',
            'detail1_name' : '東京五輪 IOC 一行接待\n（2021年7月）',
            'count1'       : '1 セット',
            'price1'       : '50,000',
            'amount1'      : '50,000',
            'detail2_name' : '東京五輪 IOC 一行ファミリ宿泊代',
            'count2'       : '2000 人',
            'price2'       : '100',
            'amount2'      : '200,000',
            'detail3_name' : '',
            'count3'       : '',
            'price3'       : '',
            'amount3'      : '',
            'detail4_name' : '',
            'count4'       : '',
            'price4'       : '',
            'amount4'      : '',
            'name2'        : '都民ネクスト',
            'address1'     : '〒xxxx-xxxx 東京都新宿区',
            'tel1'         : 'TEL 030-1111-0000',
            'bank1'        : 'メガバンク',
            'bank1_1'      : '新宿中央店',
            'folio1'       : '秘密口座',
            'folio1_no'    : '１１１１１',
        }
    }
    makePDFwithExcel(mkInfo)

#[EOF]