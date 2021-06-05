# -*- coding: utf-8 -*-

import os

#使用するフォント
if os.name == 'posix':
    # for linux / android Termux
    _FontList = {
        'Arial': '/data/data/com.termux/files/home/.fonts/IPAexfont00401/ipaexg.ttf',
        'ＭＳ ゴシック': '/data/data/com.termux/files/home/.fonts/IPAexfont00401/ipaexg.ttf',
        'ＭＳ Ｐゴシック': '/data/data/com.termux/files/home/.fonts/IPAexfont00401/ipaexg.ttf',
    }
elif os.name == 'nt':
    _FontList = {
        'Arial': 'C:\\Windows\\Fonts\\arial.ttf',
        'ＭＳ ゴシック': 'C:\\Windows\\Fonts\\msgothic.ttc',
        'ＭＳ Ｐゴシック': 'C:\\Windows\\Fonts\\msgothic.ttc',
}

#[EOF]