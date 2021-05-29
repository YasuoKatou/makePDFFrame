# -*- coding: utf-8 -*-

import pathlib
import re

import openpyxl
import reportlab
from reportlab.pdfbase import pdfmetrics as RPPdfmetrics
from reportlab.pdfbase import ttfonts as RPTTFonts
from reportlab.pdfgen import canvas as RPCanvas
from reportlab.lib import pagesizes as PLPageSize
from reportlab.lib.units import inch, mm
from reportlab.platypus import Table as RPTable
from reportlab.platypus import TableStyle as RPTableStyle
from reportlab.lib import colors as RLColors

import Const
import Excel2Json

_tempWB = openpyxl.Workbook()
_tempWS = _tempWB.active
_WSJson = None

for n, p in Const._FontList.items():
    RPPdfmetrics.registerFont(RPTTFonts.TTFont(n, p))

#インチは25.4mm （ミリメートル）
INCH2MM = 25.4
#1ポイントは1／72インチ
DPI = 72.0
_columnWidths = []
_rowHeights = []
_leftPos = []
_topPos  = []
_sizeUnit = mm
#サイズは、Excelより取得したいが、正しく取得できないので引数とする
#また、単位を「mm」とする
#TODO 将来は起動時引数とする
_args = None

_excelDataValiableFmt = re.compile(r'\{\#(?P<valiable_name>\S+)\}')

def _sizeInit(wsJson):
    pm = wsJson['page_margins']
    # 左余白
    if pm['left'] != None:
        leftMargin = pm['left']
        #msgFmt = '[page margin left: %finch(%fmm)]'
        msgFmt = '<page margin left: %fmm>'
    else:
        leftMargin = 0.0
        msgFmt = '<page margin left: %finch(%fmm)>'
    #print(msgFmt % (leftMargin, leftMargin * INCH2MM))
    print(msgFmt % (leftMargin, ))
    # 上余白
    if pm['top'] != None:
        topMargin = pm['top']
        #msgFmt = '[page margin top : %finch(%fmm)]'
        msgFmt = '<page margin top : %ffmm>'
    else:
        topMargin = 0.0
        msgFmt = '<page margin top : %finch(%fmm)>'
    #print(msgFmt % (topMargin, topMargin * INCH2MM))
    print(msgFmt % (topMargin, ))
    # 列幅の一覧（印刷領域）
    defaultSize = openpyxl.worksheet.dimensions.SheetFormatProperties()
    left = leftMargin
    _leftPos.append(left)
    for w in _WSJson['row_height']['width'].values():
        #pos = w if w is not None else defaultSize.baseColWidth
        pos = wsJson['row_height']['width']
        _columnWidths.append(pos)
        left += pos
        _leftPos.append(left)
    # 行高の一覧（印刷領域）
    top = topMargin
    _topPos.append(top)
    for h in _WSJson['row_height']['height'].values():
        #pos = h if h is not None else defaultSize.defaultRowHeight
        pos = wsJson['row_height']['height']
        _rowHeights.append(pos)
        top += pos
        _topPos.append(top)

def _pdfAttribute(pdf_canvas):
    pdf_canvas.saveState()
    pdf_canvas.setAuthor("Y.Katou (YKS)") # 作者
    pdf_canvas.setTitle("テスト") # 表題
    pdf_canvas.setSubject("preview") # 件名

def _newPdf(pdfPath):
    printPageSetup = _WSJson['PrintPageSetup']
    # 用紙サイズ
    paperSize = printPageSetup['paperSize']
    print('[paperSize] : %s' % paperSize)
    if not paperSize:
        paperSize = _tempWS.PAPERSIZE_A4
    if printPageSetup['paperSize'] != paperSize:
        print('<paperSize> : %s' % paperSize)
    # 用紙サイズは、A4のみ対応する
    assert str(paperSize) == _tempWS.PAPERSIZE_A4, 'page size not suportted.'
    paperSize = reportlab.lib.pagesizes.A4
    # 用紙の方向
    orientation = printPageSetup['orientation']
    print('[orientation] : %s' % orientation)
    if orientation != _tempWS.ORIENTATION_LANDSCAPE:    # 横置き以外
        orientation = _tempWS.ORIENTATION_PORTRAIT
    if printPageSetup['orientation'] != orientation:
        print('<orientation> : %s' % orientation)
    # PDFオブジェクトを生成
    if orientation == _tempWS.ORIENTATION_LANDSCAPE:
        pagesize = PLPageSize.landscape(paperSize)
    else:
        pagesize = PLPageSize.portrait(paperSize)
    # 左上が原点（bottomup=False）
    pdf_canvas = RPCanvas.Canvas(pdfPath, pagesize=paperSize)
    pdf_canvas.setFont("Times-Roman", 10)
    return pdf_canvas

def _getCellRect(a1, isStr=True):
    printArea = openpyxl.utils.range_boundaries(_WSJson['print_area'])
    #print(printArea)
    cellAddress = openpyxl.utils.range_boundaries(a1)
    #if isStr:
    #    # 文字列は、１行下に座標を下げる
    #    cellAddress = (
    #        cellAddress[0],
    #        cellAddress[1] + 1,
    #        cellAddress[2],
    #        cellAddress[3] + 1,
    #    )
    #print(cellAddress)
    return {
        'left'  : _leftPos[cellAddress[0] - printArea[0]],
        #'top'   : _topPos[cellAddress[1]  - printArea[1]]     / DPI,
        'top'   : _topPos[cellAddress[1]  - printArea[1]],
        'right' : _leftPos[cellAddress[2] - printArea[0] + 1],
        #'bottom': _topPos[cellAddress[3]  - printArea[1] + 1] / DPI,
        'bottom': _topPos[cellAddress[3]  - printArea[1] + 1],
    }

def _drawString2(pdf_canvas, cell, valiableData=None):
    def __getString(s):
        if not valiableData:
            return s
        m = re.match(_excelDataValiableFmt, s)
        if m:
            n = m.group('valiable_name')
            if n in valiableData:
                return valiableData[n]
        return s
    print('[A1: %s, String: %s, Font:(%s, %d), %s, %s]' %
        (cell['A1'], cell['value'], cell['font']['name'], cell['font']['size']
        , cell['alignment']['horizontal'], cell['alignment']['vertical']))
    cellString = __getString(cell['value'])
    print(cellString)
    # フォント
    styles = [('FONT', (0, 0), (-1, -1), cell['font']['name'], cell['font']['size'])]
    if _debug:
        styles.append(('BOX', (0, 0), (-1, -1), 1, RLColors.red))
    # 横方向の位置
    if cell['alignment']['horizontal'] == 'right':
        styles.append(('ALIGN', (0, 0), (-1, -1), 'RIGHT'))
    elif cell['alignment']['horizontal'] == 'center':
        styles.append(('ALIGN', (0, 0), (-1, -1), 'CENTER'))
    else:
        styles.append(('ALIGN', (0, 0), (-1, -1), 'LEFT'))
    # 縦方向の位置
    if cell['alignment']['vertical'] == 'right':
        styles.append(('VALIGN', (0, 0), (-1, -1), 'TOP'))
    elif cell['alignment']['vertical'] == 'center':
        styles.append(('VALIGN', (0, 0), (-1, -1), 'MIDDLE'))
    elif cell['alignment']['vertical'] == 'bottom':
        styles.append(('VALIGN', (0, 0), (-1, -1), 'BOTTOM'))

    cellRect = _getCellRect(cell['A1'])
    h = cellRect['bottom'] - cellRect['top']
    t = RPTable([[cellString]]
              , colWidths=(cellRect['right']  - cellRect['left']) * _sizeUnit
              , rowHeights=(h                                     * _sizeUnit)
              , style=RPTableStyle(styles))
    x = cellRect['left'] * _sizeUnit
    y = (297.0 - h - cellRect['top'])  * _sizeUnit
    t.wrapOn(pdf_canvas, x, y)
    t.drawOn(pdf_canvas, x, y)

def _drawString(pdf_canvas, cell):
    print('[A1: %s, String: %s, Font:(%s, %d)]' %
        (cell['A1'], cell['value'], cell['font']['name'], cell['font']['size'], ))
    cellRect = _getCellRect(cell['A1'])
    print(cellRect)
    pdf_canvas.setFont(cell['font']['name'], cell['font']['size'])
    #pdf_canvas.drawString(cellRect['left'] * inch, cellRect['top'] * inch, 'String')
    if cell['alignment']['horizontal'] == 'center':
        x = (cellRect['right'] - cellRect['left']) / 2 + cellRect['left']
        pdf_canvas.drawCentredString(x               * _sizeUnit
                                   , cellRect['top'] * _sizeUnit
                                   , str(cell['value']))
    elif cell['alignment']['horizontal'] == 'right':
        pdf_canvas.drawRightString(cellRect['right'] * _sizeUnit
                                 , cellRect['top']   * _sizeUnit
                                 , str(cell['value']))
    else:
        pdf_canvas.drawString(cellRect['left'] * _sizeUnit
                            , cellRect['top']  * _sizeUnit
                            , str(cell['value']))

def _drawBoarders(c, boarders):
    for boarder in boarders:
        r = _getCellRect(boarder['A1'], isStr=False)
        print('[%s] - %s' % (boarder['A1'], str(r)))
        if boarder['kind'] == Excel2Json._BOARDER_TYPE.BOX_LEFT_TOP:
            h = (r['bottom'] - r['top'])
            y = 297.0 - h - r['top']
            c.rect(r['left'] * _sizeUnit
                 , y         * _sizeUnit
                 , (r['right'] - r['left']) * _sizeUnit
                 , h                        * _sizeUnit
                 , stroke=1, fill=0)
        elif boarder['kind'] == Excel2Json._BOARDER_TYPE.BOTTOM_ONLY:
            y = 297.0 - r['bottom']
            c.line(r['left']   * _sizeUnit
                 , y           * _sizeUnit
                 , r['right']  * _sizeUnit
                 , y           * _sizeUnit)

def _drawGrid(c):
    c.setStrokeColor(RLColors.lavender)
    # A4サイズ. 210 × 297 ミリ
    #縦線
    for x in range(0, 21):
        c.line(x * 10.0 * mm
             , 0.0      * mm
             , x * 10.0 * mm
             , 297      * mm)
    #横線
    for y in range(0, 30):
        c.line(0.0      * mm
             , y * 10.0 * mm
             , 210.0    * mm
             , y * 10.0 * mm)

def makePDFwithExcel(mkInfo):
    global _debug, _WSJson, _args
    _args = mkInfo['args']
    _debug = True if mkInfo['_debug'] else False

    wbJson = Excel2Json.readExcel(mkInfo['ExcelPath'])
    if _debug:
        print(wbJson)

    pdfPath = mkInfo['PDFPath']
    if not pdfPath:
        p = pathlib.Path(mkInfo['ExcelPath'])
        pdfPath = str(p.parent / p.stem) + '.pdf'
    for sn, _WSJson in wbJson.items():

        if sn != mkInfo['SheetName']:
            continue
        #_sizeInit(_WSJson)
        _sizeInit(_args)
        pdf_canvas = _newPdf(pdfPath)
        _pdfAttribute(pdf_canvas)
        #pdf_canvas.drawString(0.1 * inch, 0.1 * inch, 'Origin')
        for cell in _WSJson['cells']:
            _drawString2(pdf_canvas, cell, mkInfo['data'])
        if _WSJson['boarders']:
            _drawBoarders(pdf_canvas, _WSJson['boarders'])
        if _debug:
            _drawGrid(pdf_canvas)
    pdf_canvas.showPage()
    pdf_canvas.save() # 保存

if __name__ == '__main__':
    excelFile = 'V01-frame_001_LibreOffice.xlsx'
    excelPath = pathlib.Path(__file__).parent / 'FrameExcel' / excelFile
    mkInfo = {
        'ExcelPath': str(excelPath),
        'SheetName': 'Sheet3',
        'PDFPath'  : '',
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
        'data'     : None
    }
    makePDFwithExcel(mkInfo)

#[EOF]