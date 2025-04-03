'''
@author: tjh
@date: from 2022-09 to 2024-07
@version: 0.2.0

changes:
1.日期参数可为ISOformat格式字符串
2.增加容错功能代码
'''


from datetime import date, timedelta, datetime
from string import ascii_uppercase


def backwardjxr(thisdate: 'datetime | date | str', basedate, months: int):
    '''获取当前日期前最近的结息基准日

example:    
thisdate, basedate, months = date(2022,5,5), date(2020,1,31), 3
thisdate、basedate可为'yyyy-mm-dd'或'yyyymmdd'格式字符串
backwardjxr(thisdate, basedate, months)->datetime.date(2022, 4, 30)
类似EXCEL公式: 
    '=EDATE(basedate,INT(DATEDIF(basedate,thisdate,"m")/months)*months)
'''

    try:
        if isinstance(thisdate, str):
            thisdate = date.fromisoformat(thisdate)
        if isinstance(basedate, str):
            basedate = date.fromisoformat(basedate)

        b_year, b_month, b_day = basedate.year, basedate.month, basedate.day
        i = (thisdate.year - b_year)*12 + thisdate.month - b_month
        if thisdate.day <= b_day:
            i -= 1
        dmonths = int((i//months)*months)
        n_year, n_month, n_day = b_year + (b_month+dmonths-1)//12, \
            (b_month+dmonths-1) % 12+1, b_day
        if n_year % 4 != 0 or (n_year % 100 == 0 and n_year % 400 != 0):
            lm = [-1, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        else:
            lm = [-1, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if n_day > lm[n_month]:
            n_day = lm[n_month]
        if isinstance(thisdate, datetime):
            return datetime(n_year, n_month, n_day)
        else:
            return date(n_year, n_month, n_day)
    except Exception as e:
        print(e)
        return None


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


edate_months = edate


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


def vlookup(lookup_value, table_array, col_index_num, range_lookup=True):
    '''vlookup(lookup_value, table_array, col_index_num)
    模拟EXCEL的Vlookup函数，模糊查询需要table_array升序排序
    不同的是的栏次以0开始，
    '''

    result = None
    n = len(table_array)
    if range_lookup:
        for i in range(-1, -n-1, -1):
            if table_array[i][0] <= lookup_value:
                result = table_array[i][col_index_num]
                break
    else:
        for i in range(n):
            if table_array[i][0] == lookup_value:
                result = table_array[i][col_index_num]
                break
    return result
