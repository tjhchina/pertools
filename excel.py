from datetime import date, timedelta, datetime
from string import ascii_uppercase
import xlwings as xw


def get_notes(fzx: str = '线别', color: tuple = (204, 255, 204), ws: xw.Sheet = None) -> list:
    ''' 获取指定工作表中特定辅助项目并标记颜色

    Parameters:
        fzx: 辅助项目，默认为”线别“
        color:  标记颜色，默认为淡绿色(204, 255, 204)
        ws: 指定工作表对象，默认为活动工作表

    Return:
        符合条件的辅助单元格的地址及注释内容, 无符合条件时返回None

    '''

    if not ws:
        ws = xw.sheets.active
    selected_notes = None
    with ws.book.app.properties(display_alerts=False,  screen_updating=False):
        values = ws.used_range.value
        if values:
            selected_notes = []
            for r, r_value in enumerate(values):
                for c, c_value in enumerate(r_value):
                    if c_value == fzx:
                        e = ws.cells[r, c]
                        if e.note:
                            e.color = color
                            selected_notes.append(
                                (e.address, e.note.text.replace('\n', '|'))
                            )
    return selected_notes


def edate(basedate: 'datetime | date | str', months: int):
    '''获取移动相应月数的对应日期

    类似EXCEL公式: 
        EDATE(basedate,months)
        basedate可为'yyyy-mm-dd'或'yyyymmdd'格式字符串
    '''

    try:
        if isinstance(basedate, str):
            basedate = date.fromisoformat(basedate)

        y0, m0, d0 = basedate.year, basedate.month, basedate.day
        y, m = divmod(m0 + int(months), 12)
        if m != 0:
            y1, m1 = y0+y, m
        else:
            y1, m1 = y0+y-1, 12
        if y1 % 4 != 0 or (y1 % 100 == 0 and y1 % 400 != 0):
            lm = [-1, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        else:
            lm = [-1, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if d0 > lm[m1]:
            d0 = lm[m1]
        if isinstance(basedate, datetime):
            result = datetime(y1, m1, d0)
        else:
            result = date(y1, m1, d0)
        return result
    except Exception as e:
        print(e)
        return None


def edate_days(basedate: 'datetime | date | str', days: int):
    ''' 获取移动相应天数的对应日期

    basedate可为'yyyy-mm-dd'或'yyyymmdd'格式字符串
    '''

    try:
        if isinstance(basedate, str):
            basedate = date.fromisoformat(basedate)
        return basedate + timedelta(days=days)
    except Exception as e:
        print(e)
        return None


def eomonth(basedate: 'datetime | date | str', months: int):
    '''获取移动相应月数的月末日期

    类似EXCEL公式: 
        EoMonth(basedate,months)
        basedate可为'yyyy-mm-dd'或'yyyymmdd'格式字符串
    '''
    try:
        if isinstance(basedate, str):
            basedate = date.fromisoformat(basedate)

        y0, m0, d0 = basedate.year, basedate.month, basedate.day
        y, m = divmod(m0 + months, 12)
        if m != 0:
            y1, m1 = y0+y, m
        else:
            y1, m1 = y0+y-1, 12
        if y1 % 4 != 0 or (y1 % 100 == 0 and y1 % 400 != 0):
            lm = [-1, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        else:
            lm = [-1, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if isinstance(basedate, datetime):
            result = datetime(y1, m1, lm[m1])
        else:
            result = date(y1, m1, lm[m1])
        return result
    except Exception as e:
        print(e)
        return None


def get_col_name(n: int) -> str:
    '''将Excel数字形式列位置转换为Excel对应列名

    Example:
    >>>get_col_name(1), get_col_name(28) # 'A', 'AB'
    '''

    a, b = divmod(n-1, 26)
    return (ascii_uppercase[a-1] if a > 0 else '') + ascii_uppercase[b]


def colnum2name(n):
    '''Translate a colnum number to name(e.g. 1 ->'A', etc.).'''
    assert n > 0
    s = ''
    while n:
        n, m = divmod(n-1, 26)
        s = chr(m+ord('A')) + s
    return s


def colname2num(s):
    '''Translate a colnum name to number (e.g. 'A' ->1, 'AA' ->27).'''
    s = s.upper()
    n = 0
    for c in s:
        assert 'A' <= c <= 'Z'
        n = n*26 + ord(c) - ord('A') + 1
    return n

def comment_font_set(size=9, bold=False):
    '''活动工作表注释字体设置'''

    ws = xw.sheets.active
    with ws.book.app.properties(screen_updating = False):
        try:
            for c in xw.sheets.active.api.Comments:
                try:
                    font = c.Shape.TextFrame.Characters().Font
                    font.Size = size
                    font.Bold = bold
                except:
                    print(c.address)
                    continue
        except:
            pass
