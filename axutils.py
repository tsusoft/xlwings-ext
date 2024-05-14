"""
xlwings-ext

Extended xlwing utils for operating excel applications

Copyright (C) 2024-present, tsusoft.net
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""

import xlwings as xw
import time
import os
import re
import hashlib
import logging
import traceback
import sys
import platform
import subprocess

from datetime import date
from appscript import k as kw

__SUFFIX_FORMAT_TIME__ = '%Y%m%d-%H%M%S'
__SUFFIX_FORMAT_DATE__ = '%Y%m%d'

SUFFIX_SEPARATOR = '_'
STRING_WRAPPER = '"'
PLACEHOLDER_LASTDATE = '{LASTDATE}'
PARSER_ARG_ESCAPE = '@'
PARSER_CMD_REGEX = '\[.+\]$'  # '\[[0-9]+\]$'

LOG_LINE_SEPARATOR_X = '- - - - - - - - - - - - - - - -'
LOG_LINE_SEPARATOR_Y = '-------------------------------'


# ============================================================
# --
# -- Customized namespace object
# --
# ============================================================

class DictNamespace:
    def __init__(self, name, dic, parent=None):
        self._name = '' if name is None else name
        self._dic = {} if dic is None else dic
        self._parent = parent

    def __getattr__(self, attr):
        value = self._dic.get(attr)
        if value is None:
            return None

        if type(value) is dict:
            return DictNamespace(attr, value, parent=self)
        elif isinstance(value, DictNamespace):
            # value._parent = self
            return value
        else:
            return value

    def value(self, key, default=None):
        v = self.__getattr__(key)
        return default if v is None else v

    def get(self, key, default=None):
        return self.value(key, default=default)

    def put(self, key, value):
        self._dic[key] = value
        return self

    def pop(self, key, default=None):
        return self._dic.pop(key, default)

    def namespace(self, key):
        v = self.__getattr__(key)

        # Build namespace if value is empty
        if v is None:
            v = DictNamespace(key, {}, parent=self)
            self._dic[key] = v._dic
        return v

    @property
    def namespaces(self):
        # TODO Should value with type DictNamespace be considered as part of them?
        # Build and return all namespaces in the path
        return [DictNamespace(key, value, parent=self) \
                for key, value in self._dic.items() \
                if (not value is None) and (type(value) is dict)]

    @property
    def parent(self):
        return self._parent

    @property
    def top(self):
        t = self
        p = t._parent
        while not p is None:
            t = p
            p = p._parent
        return t

    def top(self, level):
        lst = [self]
        p = self._parent
        while not p is None:
            lst.append(p)
            p = p._parent

        # zero position as first position
        level = 1 if level == 0 else level
        level = -level - 1 if level < 0 else -level
        return None if abs(level) > len(lst) else lst[level]

    @property
    def namespace_names(self):
        nsl = []
        for item in self._dic.items():
            if (not item[1] is None) and (type(item[1]) is dict):
                nsl.append(item[0])
        return nsl

    @property
    def names(self):
        return self._dic.keys()

    @property
    def name(self):
        return self._name

    @property
    def dict(self):
        return self._dic

    def __str__(self):
        return "{0}(name={1}, parent={2}, dict={3})" \
            .format(self.__class__.__name__, self._name,
                    None if self._parent is None else self._parent._name,
                    self._dic.keys())

    def __repr__(self):
        return self.__str__()


# ============================================================
# --
# -- Toolkits: File manupulation
# --
# ============================================================

def raw_filename(fullname):
    return subprocess.run(['readlink', '-f', fullname], stdout=subprocess.PIPE, text=True).stdout.strip()


# ============================================================
# --
# -- Toolkits: String manupulation
# --
# ============================================================

def placeholders(str, content=False, placeholder=['{', '}']):
    regex = ''.join([placeholder[0],
                     '(.*?)' if content else '.*?',
                     placeholder[1]])
    return re.findall(regex, str)


def remove_unprintable_whitespace(s):
    # remve whitespace and blank
    cls = s.strip().replace(' ', '')
    return ''.join(x for x in cls if x.isprintable())


def prefix(name, separator='_'):
    # Return the suffix part
    if not separator is None:
        lst = name.split(separator, 1)
        return '' if len(lst) == 1 else lst[0]
    elif separator == '':
        return name
    else:
        return ''


def prefix_strip(name, prefix=None, separator='_'):
    if prefix is None:
        # Remove suffix from the name part
        if not separator is None:
            lst = name.split(separator, 1)
            name = lst[0] if len(lst) == 1 else lst[1]
    else:
        # Build prefix with separator
        prefix = prefix if separator is None else separator.join([prefix, ''])
        # Remove suffix from the name part
        name = name.removeprefix(prefix)
    return name


def suffix(name, separator='_'):
    # Return the suffix part
    if not separator is None:
        lst = name.rsplit(separator, 1)
        return '' if len(lst) == 1 else lst[1]
    elif separator == '':
        return name
    else:
        return ''


def suffix_strip(name, suffix=None, separator='_'):
    if suffix is None:
        # Remove suffix from the name part
        if not separator is None:
            lst = name.rsplit(separator, 1)
            name = lst[0] if len(lst) == 1 else lst[0]
    else:
        # Build suffix with separator
        suffix = suffix if separator is None else separator.join(['', suffix])
        # Remove suffix from the name part
        name = name.removesuffix(suffix)
    return name


# ============================================================
# --
# -- Toolkits: Date manupulation
# --
# ============================================================

def date_to_excel_ordinal(year, month, day):
    # Specifying offset value i.e.,
    # the date value for the date of 1900-01-00
    offset = 693594
    current = date(year, month, day)

    # Calling the toordinal() function to get
    # the excel serial date number in the form
    # of date values
    n = current.toordinal()
    return (n - offset)


def suffix_stamp(timesec=None, with_time=False):
    format = __SUFFIX_FORMAT_TIME__ if with_time else __SUFFIX_FORMAT_DATE__
    return time.strftime(format, time.localtime(time.time() if timesec is None else timesec))


def suffix_time(time_string, with_time=False):
    try:
        format = __SUFFIX_FORMAT_TIME__ if with_time else __SUFFIX_FORMAT_DATE__
        return time.strptime(time_string, format)
    except:
        # Try again with time if 'with_time' is given as None, which means to try both
        return suffix_time(time_string, with_time=True) if with_time is None else None


# ============================================================
# --
# -- Toolkits: xlwings manupulation
# --
# ============================================================

# ----------------------------------------
# Excel: Description objects
# ----------------------------------------

class RangeDesc:

    def __init__(self):
        # self._worksheet = None
        # self._range = None
        self._wb_name_ = None
        self._sh_name_ = None
        self._sync = False

        self._row_start = 0
        self._col_start = 0
        self._row_end = 0
        self._col_end = 0

    def __check__(self, row_start=None, col_start=None, row_end=None, col_end=None):
        r1 = self._row_start if row_start is None else row_start
        r2 = self._col_start if col_start is None else col_start
        r3 = self._row_end if row_end is None else row_end
        r4 = self._col_end if col_end is None else col_end
        if (r1 > r3) or (r2 > r4):
            raise ValueError("Not a valid range: row_s={}, col_s={}; row_e={}, col_e={}".format(r1, r2, r3, r4))
        self._row_start = r1
        self._col_start = r2
        self._row_end = r3
        self._col_end = r4

    @property
    def worksheet(self):
        # return self._worksheet
        if (not self._wb_name_ is None) and (not self._sh_name_ is None):
            return xw.Book(self._wb_name_).sheets[self._sh_name_]
        else:
            return None

    @property
    def book_name(self):
        return self._wb_name_

    @property
    def sheet_name(self):
        return self._sh_name_

    @property
    def range(self):
        # if (self._range is None) and (not self._worksheet is None):
        #    try:
        #        self._range = self._worksheet.range(self.px, self.py)
        #    except:
        #        self._range = None # Not a valid desc for excel
        # return self._range
        sh = self.worksheet
        if not sh is None:
            try:
                return sh.range(self.px, self.py)
            except:
                pass
        return None

    @property
    def row_start(self):
        return self._row_start

    @property
    def col_start(self):
        return self._col_start

    @property
    def row_end(self):
        return self._row_end

    @property
    def col_end(self):
        return self._col_end

    @property
    def top(self):
        return self._row_start

    @property
    def left(self):
        return self._col_start

    @property
    def bottom(self):
        return self._row_end

    @property
    def right(self):
        return self._col_end

    @property
    def in_range(self, pt):
        if pt is None:
            return False
        return (pt[0] >= self._row_start) \
            and (pt[0] <= self._row_end) \
            and (pt[1] >= self._col_start) \
            and (pt[1] <= self._col_end)

    @property
    def rangeable(self):
        return self._row_start > 0 and self._row_end > 0 and self._col_start > 0 and self._col_end > 0

    def intersect(self, target):
        _worksheet = self.worksheet
        if (target is None) or (target.worksheet != _worksheet):
            return None

        r1, r2 = self.px
        r3, r4 = self.py

        # Target not in scope
        if (target.top > self.bottom) \
                or (target.bottom < self.top) \
                or (target.left > self.right) \
                or (target.right < self.left):
            return None

        if (target.top >= self.top) and (target.top <= self.bottom):
            r1 = target._row_start
        if (target.left >= self.left) and (target.left <= self.right):
            r2 = target._col_start
        if (target.bottom >= self.top) and (target.bottom <= self.bottom):
            r3 = target._row_end
        if (target.right >= self.left) and (target.right <= self.right):
            r4 = target._col_end

        r = RangeDesc().attach(_worksheet)
        return r.update(row_start=r1,
                        col_start=r2,
                        row_end=r3,
                        col_end=r4)

    def shift_away(self, target, mode='down', entire=True):
        if (self.intersect(target) is None):
            return self

        gap_h = target.right - self.left
        gap_v = target.bottom - self.top
        debug('RangeDesc::shift_away', 'target.right={}, self.left={}, gap_h={}', target.right, self.left, gap_h)
        debug('RangeDesc::shift_away', 'target.bottom={}, self.top={}, gap_v={}', target.bottom, self.top, gap_v)

        if mode == 'right':
            rng = None
            if entire:
                rng = target.worksheet.range('{}:{}'.format(
                    dec2alphabet(self.left),
                    dec2alphabet(self.left + gap_h)))
            else:
                rng = target.worksheet.range(self.px,
                                             (self._row_end, self.left + gap_h))
            debug('RangeDesc::shift_away', 'mode={}, rng={}', mode, rng)
            rng.insert(shift=mode)
            return self.move((self._row_start, self.left + gap_h + 1))
        elif mode == 'down':
            rng = None
            if entire:
                rng = target.worksheet.range('{}:{}'.format(
                    self.top, self.top + gap_v))
            else:
                rng = target.worksheet.range(self.px,
                                             (self.top + gap_v, self._col_end))
            debug('RangeDesc::shift_away', 'mode={}, rng={}', mode, rng)
            rng.insert(shift=mode)
            return self.move((self.top + gap_v + 1, self._col_start))
        else:
            raise ValueError('Unknown shift mode {}'.format(mode))

    def expand(self, mode='table'):
        rng = self.range
        if rng is None:
            return self

        # Don't expand if empty cell
        if (rng[0, 0].value is None) or (rng[0, 0].value == ''):
            return self

        return self.update_by(rng.expand(mode=mode).address)

    def offset(self, row_offset=None, column_offset=None):
        vr = 0 if row_offset is None else row_offset
        vc = 0 if column_offset is None else column_offset
        self.update(row_start=self._row_start + vr,
                    col_start=self._col_start + vc,
                    row_end=self._row_end + vr,
                    col_end=self._col_end + vc)
        return self

    def resize(self, row_size=None, column_size=None):
        vr, vc = 0, 0
        if not row_size is None:
            vr = row_size - self.height

        if not column_size is None:
            vc = column_size - self.width

        self.update(row_end=self._row_end + vr,
                    col_end=self._col_end + vc)
        return self

    def move(self, px):
        t = type(px)

        if t is tuple:
            vr = px[0] - self._row_start
            vc = px[1] - self._col_start
            self.update(row_start=px[0],
                        col_start=px[1],
                        row_end=self._row_end + vr,
                        col_end=self._col_end + vc)
        elif t is xw.Range:
            rg = px
            vr = rg.row - self._row_start
            vc = rg.column - self._col_start
            self.update(row_start=rg.row,
                        col_start=rg.column,
                        row_end=self._row_end + vr,
                        col_end=self._col_end + vc)
        elif t is str:
            try:
                rg = self.worksheet.range(px)
                vr = rg.row - self._row_start
                vc = rg.column - self._col_start
                self.update(row_start=rg.row,
                            col_start=rg.column,
                            row_end=self._row_end + vr,
                            col_end=self._col_end + vc)
            except:
                pass
        return self

    def update(self, row_start=None, col_start=None, row_end=None, col_end=None):
        self.__check__(row_start=row_start, col_start=col_start, row_end=row_end, col_end=col_end)
        #       self._row_start = self._row_start if row_start is None else row_start
        #       self._col_start = self._col_start if col_start is None else col_start
        #       self._row_end = self._row_end if row_end is None else row_end
        #       self._col_end = self._col_end if col_end is None else col_end

        #       if not self._worksheet is None:
        #           try:
        #               self._range = self._worksheet.range(self.px, self.py)
        #           except:
        #               self._range = None # Not a valid desc for excel
        return self

    def update_from(self, desc):
        if not desc is None:
            # self._worksheet = desc.worksheet
            self._wb_name_ = desc._wb_name_
            self._sh_name_ = desc._sh_name_
            self.update(row_start=desc.row_start,
                        col_start=desc.col_start,
                        row_end=desc.row_end,
                        col_end=desc.col_end)
        return self

    def update_by(self, cell1=None, cell2=None):
        _worksheet = self.worksheet
        if not _worksheet is None:
            try:
                _range = _worksheet.range(cell1, cell2)
                self._row_start = _range.row
                self._col_start = _range.column
                self._row_end = _range.last_cell.row
                self._col_end = _range.last_cell.column
            except:
                pass  # nothing happened when there is an exception #self._range = None
        return self

    def update_by_xy(self, px=None, py=None):
        r1 = self._row_start if px is None else px[0]
        r2 = self._col_start if px is None else px[1]
        r3 = self._row_end if py is None else py[0]
        r4 = self._col_end if py is None else py[1]
        self.__check__(row_start=r1, col_start=r2, row_end=r3, col_end=r4)
        #       if not self._worksheet is None:
        #           try:
        #               self._range = self._worksheet.range(self.px, self.py)
        #           except:
        #               self._range = None # Not a valid desc for excel
        return self

    @property
    def row_x(self):
        return self._row_start

    @property
    def col_x(self):
        return self._col_start

    @property
    def row_y(self):
        return self._row_end

    @property
    def col_y(self):
        return self._col_end

    @property
    def px(self):
        return (self._row_start, self._col_start)

    @property
    def py(self):
        return (self._row_end, self._col_end)

    @property
    def tl(self):
        return self.px

    @property
    def tr(self):
        return (self._row_start, self._col_end)

    @property
    def bl(self):
        return (self._row_end, self._col_start)

    @property
    def br(self):
        return self.py

    @property
    def args(self):
        return (self.px, self.py)

    @property
    def width(self):
        return self._col_end - self._col_start + 1

    @property
    def height(self):
        return self._row_end - self._row_start + 1

    def attach(self, worksheet):
        # Return current if None is given
        if worksheet is None:
            return self

        # self._worksheet = worksheet
        self._wb_name_ = worksheet.book.fullname
        self._sh_name_ = worksheet.name
        #       try:
        #           self._range = self._worksheet.range(self.px, self.py)
        #       except:
        #           self._range = None # Not a valid desc for excel
        return self

    def detach(self):
        # self._worksheet = None
        # self._range = None
        self._wb_name_ = None
        self._sh_name_ = None
        return self

    @property
    def address(self):
        _range = self.range
        return None if _range is None else _range.address

    def duplicate(self):
        cloned = RangeDesc().update(row_start=self._row_start,
                                    col_start=self._col_start,
                                    row_end=self._row_end,
                                    col_end=self._col_end)
        cloned.attach(self.worksheet)
        return cloned

    def __str__(self):
        return "{0}(args={1}, worksheet={2}, range={3})" \
            .format(self.__class__.__name__, (self.px, self.py),
                    self.worksheet, self.range)

    def __repr__(self):
        return self.__str__()


class SheetDesc:

    def __init__(self):
        self._title = RangeDesc()
        self._formulas = RangeDesc()
        self._data = RangeDesc()

    @property
    def title(self):
        return self._title

    @property
    def formulas(self):
        return self._formulas

    @property
    def data(self):
        return self._data

    def duplicate(self):
        cloned = SheetDesc()
        cloned._title = self._title.duplicate()
        cloned._formulas = self._formulas.duplicate()
        cloned._data = self._data.duplicate()
        return cloned

    def update_from(self, desc):
        if not desc is None:
            self._title.update_from(desc.title)
            self._formulas.update_from(desc.formulas)
            self._data.update_from(desc.data)
        return self

    def update(self, title=None, formulas=None, data=None):
        self._title.update_from(title)
        self._formulas.update_from(formulas)
        self._data.update_from(data)
        return self

    def __str__(self):
        return "{0}(title={1}, formulas={2}, data={3})" \
            .format(self.__class__.__name__,
                    self._title, self._formulas, self._data)


class SheetDescPair:

    def __init__(self):
        self._src = SheetDesc()
        self._dst = SheetDesc()

    @property
    def src(self):
        return self._src

    @property
    def dst(self):
        return self._dst

    def duplicate(self):
        cloned = SheetDescPair()
        cloned._src = self._src.duplicate()
        cloned._dst = self._dst.duplicate()
        return cloned

    def update_from(self, desc):
        if not desc is None:
            self._src.update_from(desc.src)
            self._dst.update_from(desc.dst)
        return self

    def update(self, src=None, dst=None):
        self._src.update_from(src)
        self._dst.update_from(dst)
        return self

    def __str__(self):
        return "{0}(src={1}, dst={2})".format(self.__class__.__name__, self._src, self._dst)


class RangeDescPair:

    def __init__(self):
        self._src = RangeDesc()
        self._dst = RangeDesc()

    def args(self):
        return (self._src.px, self._src.py, self._dst.px, self._dst.py)

    @property
    def src(self):
        return self._src

    @property
    def dst(self):
        return self._dst

    def duplicate(self):
        cloned = RangeDescPair()
        cloned._src = self._src.duplicate()
        cloned._dst = self._dst.duplicate()
        return cloned

    def update_from(self, desc):
        if not desc is None:
            self._src.update_from(desc.src)
            self._dst.update_from(desc.dst)
        return self

    def update(self, src=None, dst=None):
        self._src.update_from(src)
        self._dst.update_from(dst)
        return self

    def __str__(self):
        return "{0}(src={1}, dst={2})".format(self.__class__.__name__, self._src, self._dst)


# ----------------------------------------
# General: ranges and names
# ----------------------------------------
def silence_mode(app=None):
    if app is None:
        if len(xw.apps) == 0:
            app = xw.App()
        else:
            app = xw.apps.active
    p_states = (app.screen_updating, app.display_alerts, app.visible)

    ''' return #### TODO Dummy it
    app.screen_updating = False
    app.display_alerts = False
    app.visible = False
    # '''

    app.calculation = 'automatic'
    return p_states


def normal_mode(app=None, states=None):
    if app is None:
        if len(xw.apps) > 0:
            app = xw.apps.active
        else:
            return

    p_states = (app.screen_updating, app.display_alerts, app.visible)
    if not states is None:
        app.screen_updating, app.display_alerts, app.visible = states
    else:
        app.screen_updating = True
        app.display_alerts = True
        app.visible = True

    app.calculation = 'automatic'
    return p_states


def book(bookname, **kargs):
    log('Touch {}'.format(bookname))
    silence_mode()
    wb = xw.Book(bookname, **kargs)
    silence_mode(app=wb.app)
    return wb


def close_book(bookname, **kargs):
    # return #### TODO Dummy it
    dstbook = book(bookname, **kargs)
    dstbook.close()


def close_books():
    books = [b.fullname for b in xw.books]
    for book in books:
        close_book(book)


def close_apps():
    # return #### TODO Dummy it
    apps = [a for a in xw.apps]
    for app in apps:
        app.quit()


# ----------------------------------------

# ----------------------------------------

def find_name(workbook, worksheet, name):
    # Find in workbook if not None
    if (not workbook is None) and (name in workbook.names):
        n = workbook.names[name]
        try:
            if not worksheet is None:
                if n.refers_to_range.sheet.name == worksheet.name:
                    return n.refers_to_range
            else:
                return n.refers_to_range
        except:
            logger.warning(traceback.
                           format_exception_only(sys.exc_info()[0],
                                                 sys.exc_info()[1]))
            logger.warning(traceback.format_exc())

    # Find in worksheet if not None
    if (not worksheet is None) and (name in worksheet.names):
        return worksheet.names[name].refers_to_range

    # Find in all sheets if not given one 
    if (not workbook is None) and (worksheet is None):
        for s in workbook.sheets:
            if name in s.names:
                return s.names[name].refers_to_range

    # Return None after all tries
    return None


# ----------------------------------------
# Excel: platform-dependent api methods
# ----------------------------------------

def turn_off_filtermode(sh):
    if platform.platform()[0:3].capitalize() == "Win":
        sh.api.AutoFilterMode.Set(False)
    #       if sh.api.FilterMode():
    #           sh.api.ShowAllData()
    else:
        sh.api.autofilter_mode.set(False)


#       if sh.api.filter_mode():
#           sh.api.show_all_data()


def range_apply_sort(desc, desc_sort=None,
                     order=xw.constants.SortOrder.xlAscending,
                     orientation=xw.constants.SortOrientation.xlSortColumns):
    if platform.platform()[0:3].capitalize() == "Win":
        # TODO need Win implementation
        if desc_sort is None:
            desc.worksheet.api.SortObject.SetSortRange(rng=desc.range.api)
            desc.worksheet.api.SortObject.ApplySort()
        else:
            desc.range.api.Sort(Key1=desc_sort.range.api,
                                Order1=order,
                                Orientation=orientation)  # Set default timeout as 5 minutes
    else:
        if desc_sort is None:
            desc.worksheet.api.sort_object.set_sort_range(rng=desc.range.api)
            desc.worksheet.api.sort_object.apply_sort(timeout=600)
        else:
            desc.range.api.sort(key1=desc_sort.range.api,
                                order1=order,
                                orientation=orientation,
                                timeout=600)  # Set default timeout as 5 minutes


# ----------------------------------------
# Excel: pivot refreshing and callback
# ----------------------------------------
def refresh_worksheet_pivot(sheet, filter_pvt=None,
                            pre_callback=None,
                            post_callback=None):
    # Get pivot names in each sheet
    names = sheet.api.pivot_tables.name()
    # Get pivot count in each sheet
    count = len(names) if type(names) is list else 0

    # Iterate pivots in sheet and refresh table
    for i in range(0, count):
        pvt = sheet.api.pivot_tables[i + 1]

        # Ignore if filter-by-pivot is set and not matched
        if (not filter_pvt is None) and (not filter_pvt(pvt)):
            continue

        log(">>> processing pivot [{}]{} ...".format(sheet.name, pvt.name()))

        if not pre_callback is None:
            pre_callback(sheet, pvt)

        pvt.refresh_table(timeout=-2)  # 1800)

        if not post_callback is None:
            post_callback(sheet, pvt)


def refresh_workbook_pivot(workbook,
                           filter_sht=None, filter_pvt=None,
                           pre_callback_sht=None, post_callback_sht=None,
                           pre_callback_pvt=None, post_callback_pvt=None):
    for sh in workbook.sheets:
        # Ignore if filter-by-sheet is set and not matched
        if (not filter_sht is None) and (not filter_sht(sh)):
            continue

        log(">>> processing worksheet [{}]{} ...".format(workbook.name, sh.name))

        if not pre_callback_sht is None:
            pre_callback_sht(sh)

        refresh_worksheet_pivot(sh, filter_pvt=filter_pvt,
                                pre_callback=pre_callback_pvt,
                                post_callback=post_callback_pvt)

        if not post_callback_sht is None:
            post_callback_sht(sh)


def pivot_item_filter(sh, pvt, fld, items, item, visible):
    try:
        if items is None:
            items = fld.pivot_items.name()

        if item in items:
            fld.pivot_items[item].visible.set(visible, timeout=-2)
    except:
        logger.warning(traceback.
                       format_exception_only(sys.exc_info()[0],
                                             sys.exc_info()[1]))
        # logger.warning(traceback.format_exc())
        try:
            # If exception occurred (cannot get reference)
            # Then assume the pivot has a only item that cannot be invisible
            # Then we use label filter instead
            label_filter = xw.Book('macro.xlsm').macro('PivotLabelFilter')
            label_filter(sh.book.name, sh.name, pvt.name(), fld.name(),
                         xw.constants.PivotFilterType.xlCaptionEquals if visible else \
                             xw.constants.PivotFilterType.xlCaptionDoesNotEqual, item)
        except:
            logger.warning(traceback.
                           format_exception_only(sys.exc_info()[0],
                                                 sys.exc_info()[1]))
            # logger.warning(traceback.format_exc())


# ----------------------------------------
# Excel: fundamental copy and paste
# ----------------------------------------

def hack_paste(rng, paste=None, operation=None, skip_blanks=False, transpose=False, **kargs):
    pastes = {
        # all_merging_conditional_formats unsupported on mac
        "all": kw.paste_all,
        "all_except_borders": kw.paste_all_except_borders,
        "all_using_source_theme": kw.paste_all_using_source_theme,
        "column_widths": kw.paste_column_widths,
        "comments": kw.paste_comments,
        "formats": kw.paste_formats,
        "formulas": kw.paste_formulas,
        "formulas_and_number_formats": kw.paste_formulas_and_number_formats,
        "validation": kw.paste_validation,
        "values": kw.paste_values,
        "values_and_number_formats": kw.paste_values_and_number_formats,
        None: None,
    }

    operations = {
        "add": kw.paste_special_operation_add,
        "divide": kw.paste_special_operation_divide,
        "multiply": kw.paste_special_operation_multiply,
        "subtract": kw.paste_special_operation_subtract,
        None: None,
    }

    rng.api.paste_special(
        what=pastes[paste],
        operation=operations[operation],
        skip_blanks=skip_blanks,
        transpose=transpose,
        **kargs
    )


# Copy and paste to destination
def copy_paste(src_sht, dst_sht, src_x, src_y, dst_x, dst_y):
    if dst_sht is None:
        dst_sht = src_sht
    s_rng = src_sht.range(src_x, src_y).options(chunksize=10_000)
    d_rng = dst_sht.range(dst_x, dst_y).options(chunksize=10_000)
    # s_rng.copy(destination = d_rng)
    # Hacking to do copy by using api
    s_rng.api.copy_range(destination=d_rng.api, timeout=600)


# Copy and paste values on selection, replace the original data
def copy_paste_self_v(sht, x, y):
    rng = sht.range(x, y).options(chunksize=10_000)
    rng.copy()
    # rng.paste(paste='values')
    # Hacking to paste by using api
    hack_paste(rng, paste='values', timeout=600)


# Copy to destination, and copy-paste values in destination (replace)
def copy_paste_v(src_sht, dst_sht, src_x, src_y, dst_x, dst_y):
    copy_paste(src_sht, dst_sht, src_x, src_y, dst_x, dst_y)
    copy_paste_self_v(dst_sht, dst_x, dst_y)


# Copy and paste values, and then paste formats
def copy_paste_vf(src_sht, dst_sht, src_x, src_y, dst_x, dst_y, *, with_format=False):
    s_rng = src_sht.range(src_x, src_y).options(chunksize=10_000)
    d_rng = dst_sht.range(dst_x, dst_y).options(chunksize=10_000)
    s_rng.copy()
    # d_rng.paste(paste='values')
    # Hacking to paste by using api
    hack_paste(d_rng, paste='values', timeout=600)
    if with_format:
        # d_rng.paste(paste='formats')
        # Hacking to paste by using api
        hack_paste(d_rng, paste='formats', timeout=600)


# ============================================================
# --
# -- Miscellaneous: tools
# --
# ============================================================

def attr(ns, at):
    return getattr(ns, at) if hasattr(ns, at) else None


def alphabet2dec(alphabet):
    alphabet = alphabet.upper()
    if alphabet == '' or len(alphabet) == 0:
        return 0
    if len(alphabet) == 1:
        return ord(alphabet) - 64
    else:
        return alphabet2dec(alphabet[1:]) + \
            (26 ** (len(alphabet) - 1)) * \
            (ord(alphabet[0]) - 64)


def dec2alphabet(dec):
    if dec <= 0:
        return ''
    elif dec <= 26:
        return chr(64 + dec)
    else:
        return dec2alphabet(int((dec - 1) / 26)) + chr(65 + (dec - 1) % 26)


def formulas_md5(wb_name, sh_name, rx, ry, separator=''):
    wb = book(wb_name)
    sh = wb.sheets[sh_name]
    fm = sh.range(rx, ry).formula
    fstr = separator.join(*fm)
    md = hashlib.md5()
    md.update(fstr.encode('utf-8'))
    return md.hexdigest()


# ============================================================
# --
# -- Toolkits: logging manupulation
# --
# ============================================================

def _msg_(module, message, *args, flag='    ', no_format=False):
    if no_format:
        return ' '.join([flag, '|{}:'.format(module), message, *args])
    else:
        return ' '.join([flag, '|{}:'.format(module),
                         message if len(args) == 0 else message.format(*args)])


def log(message, next=False):
    logger.info('{0:>2} - {1}:'.format(inc.next if next else '', message))


def warn(module, message, *args, no_format=False):
    logger.warn(_msg_(module, message, *args, flag='   >', no_format=no_format))


def info(module, message, *args, no_format=False):
    logger.info(_msg_(module, message, *args, flag='  >>', no_format=no_format))


def debug(module, message, *args, no_format=False):
    logger.debug(_msg_(module, message, *args, flag=' >>>', no_format=no_format))


def buggy(module, message, *args, no_format=False):
    logger.log(5, _msg_(module, message, *args, flag='>>>>', no_format=no_format))


def get_logger(logname):
    FORMAT = '%(asctime)s [%(levelname)-5s] %(message)s'
    logging.basicConfig(format=FORMAT)
    logging.addLevelName(5, 'BUGGY')
    log = logging.getLogger(logname)
    lvl = os.getenv('LOGGING_LEVEL')
    log.setLevel(logging.INFO if lvl is None else lvl)
    return log


class Incrementor:
    def __init__(self, n):
        self._data = n

    @property
    def next(self):
        self._data = self._data + 1
        return self._data

    @property
    def back(self):
        self._data = self._data - 1
        return self._data

    @property
    def current(self):
        return self._data


inc = Incrementor(0)
logger = get_logger('excellog')
