"""The two ordinary sheets behind the game: the cheat sheet and a checklist. Plain cells. The
checklist is the scoreboard: its last task has a checkbox, a text box that runs PlayDoom, and
after a game the host ticks it and writes the result into the row (hostvba.py, LogRun)."""
from xml.sax.saxutils import escape
import mkxlsm

CHECK_ROW = 10      # sheet row of the task whose checkbox starts the game; CHKROW in hostvba.py

# font=(size, bold, italic, colour); see mkxlsm.styles_xml
STYLES = [
    dict(font=(14, True, False, '000000')),                                                             # 0 checklist title
    dict(font=(9, False, True, '7F7F7F')),                                                              # 1 small grey note
    dict(font=(11, True, False, '000000'), fill='D9D9D9', border='808080', align='center,center'),      # 2 table header
    dict(font=(11, False, False, '000000'), border='808080', align='center,center'),                    # 3 centred cell
    dict(font=(11, False, False, '000000'), border='808080', align='left,center'),                      # 4 text cell
    dict(fill='000000'),                                                                                # 5 cheat sheet: background
    dict(font=(24, True, False, 'D01010'), fill='000000'),                                              # 6 big red title
    dict(font=(11, False, False, 'A8A8A8'), fill='000000'),                                             # 7 grey line
    dict(font=(11, True, False, 'D01010'), fill='000000', border='5A0000', align='left,center'),        # 8 table header
    dict(font=(12, True, False, 'FFD040'), fill='141414', border='5A0000', align='center,center'),      # 9 key
    dict(font=(11, False, False, 'E0E0E0'), fill='141414', border='5A0000', align='left,center'),       # 10 text
    dict(font=(14, True, False, 'D01010'), fill='000000'),                                              # 11 section title
    dict(font=(20, False, False, '9A9A9A'), fill='000000', align='left,center'),                        # 12 game sheet: the line under the frame
]
HINT_STYLE = 12
# Shown under the splash until macros run (the host clears it): what to do when a click does nothing.
HINTS = ['Nothing happens when you click? Macros are off. Click Enable Content in the bar above the sheet.',
         'If that bar is red: close this file, right-click it in its folder, choose Properties, tick Unblock, and open it again.']

TASKS = [('Check email', 'Done', ''),
         ('Update the tracker', 'Done', ''),
         ('Reply to "per my last email"', 'In progress', ''),
         ('Rename final_v2_FINAL.xlsx', 'Not started', ''),
         ('Attend the meeting that could have been an email', 'Done', ''),
         ('Make this spreadsheet less boring', '', '')]

CHEATS = [('8', 'God mode. Press again to switch it off.', 'IDDQD'),
          ('9', 'Every weapon, every key, full ammo and armor.', 'IDKFA'),
          ('0', 'Walk through walls. Press again to switch it off.', 'IDSPISPOPD')]

CONTROLS = [('W  /  Up', 'Forward'), ('S  /  Down', 'Back'), ('Left  /  Right', 'Turn'), ('A  /  D', 'Strafe'),
            ('Shift', 'Run'), ('Ctrl', 'Fire'), ('Space  /  E', 'Open doors, press switches'), ('1 - 7', 'Weapons'),
            ('Tab', 'Map'), ('Esc', 'Menu: new game, difficulty, sound volume'), ('P', 'Pause'),
            ('V', 'Sound on / off (Windows)'), ('Q', 'Quit, back to the spreadsheet')]

FACTS = ['This is the real DOOM engine from 1993, translated into about 80,000 lines of VBA.',
         'Every cell on the DOOM tab is one pixel: 320 across, 200 down.',
         'The game data is about a million numbers stored far to the right of the picture.',
         'No add-ins, no helper programs. One workbook.']

def _c(col, row, style, text=None):
    if text is None or text == '':
        return '<c r="%s%d" s="%d"/>' % (mkxlsm.col_name(col), row, style)
    return '<c r="%s%d" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (mkxlsm.col_name(col), row, style, escape(text))

def _n(col, row, style, number):
    return '<c r="%s%d" s="%d"><v>%d</v></c>' % (mkxlsm.col_name(col), row, style, number)

def _sheet(code_name, cols, rows, selected=False, gridlines=True, drawing=False, text_digits=None):
    """rows: {row number: (height or None, [cell xml, ...])}. text_digits: a range whose cells hold
    digits as text on purpose, so Excel does not flag them."""
    body = ''.join('<row r="%d"%s>%s</row>' % (r, ' ht="%s" customHeight="1"' % h if h else '', ''.join(cells))
                   for r, (h, cells) in sorted(rows.items()))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="%s"/>'
            '<sheetViews><sheetView%s%s workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>%s</sheetData>%s%s</worksheet>'
            % (code_name, '' if gridlines else ' showGridLines="0"', ' tabSelected="1"' if selected else '', cols, body,
               '<ignoredErrors><ignoredError sqref="%s" numberStoredAsText="1"/></ignoredErrors>' % text_digits if text_digits else '',
               '<drawing r:id="rId1"/>' if drawing else ''))

def checklist(base):
    s = lambda k: base + k
    rows = {1: (22, [_c(1, 1, s(0), 'Daily Checklist')]),
            2: (None, [_c(1, 2, s(1), 'Tick the last box to play. It keeps score.')]),
            4: (20, [_c(1, 4, s(2), '#'), _c(2, 4, s(2), 'Task'), _c(3, 4, s(2), 'Status'), _c(4, 4, s(2), 'Notes')])}
    for i, (task, status, notes) in enumerate(TASKS):
        r = 5 + i
        rows[r] = (20, [_n(1, r, s(3), i + 1), _c(2, r, s(4), task), _c(3, r, s(3), status), _c(4, r, s(4), notes)])
    assert 5 + len(TASKS) - 1 == CHECK_ROW
    cols = ('<col min="1" max="1" width="5" customWidth="1"/><col min="2" max="2" width="50" customWidth="1"/>'
            '<col min="3" max="3" width="14" customWidth="1"/><col min="4" max="4" width="64" customWidth="1"/>')
    tick = mkxlsm.text_shape('Tick', 2, CHECK_ROW - 1, 3, CHECK_ROW, '&#9744;', macro='PlayDoom', size=14, font='Segoe UI Symbol')
    return _sheet('Sheet3', cols, rows, drawing=True), mkxlsm.drawing_doc([tick])

def cheat_sheet(base):
    s = lambda k: base + k
    rows = {2: (34, [_c(2, 2, s(6), 'CHEAT SHEET')]),
            3: (None, [_c(2, 3, s(7), 'Press these while you play. Each one types the real code for you.')]),
            5: (20, [_c(2, 5, s(8), 'KEY'), _c(3, 5, s(8), 'WHAT IT DOES'), _c(4, 5, s(8), 'THE CODE IT TYPES')])}
    r = 6
    for key, what, code in CHEATS:
        rows[r] = (22, [_c(2, r, s(9), key), _c(3, r, s(10), what), _c(4, r, s(10), code)]); r += 1
    r += 1
    rows[r] = (24, [_c(2, r, s(11), 'CONTROLS')]); r += 1
    for key, what in CONTROLS:
        rows[r] = (20, [_c(2, r, s(9), key), _c(3, r, s(10), what), _c(4, r, s(10))]); r += 1
    r += 1
    rows[r] = (24, [_c(2, r, s(11), 'WHAT YOU ARE LOOKING AT')]); r += 1
    for line in FACTS:
        rows[r] = (None, [_c(2, r, s(7), line)]); r += 1
    cols = ('<col min="1" max="1" width="3" customWidth="1" style="%d"/><col min="2" max="2" width="18" customWidth="1" style="%d"/>'
            '<col min="3" max="3" width="52" customWidth="1" style="%d"/><col min="4" max="4" width="26" customWidth="1" style="%d"/>'
            '<col min="5" max="16384" width="9" style="%d"/>' % ((s(5),) * 5))
    return _sheet('Sheet2', cols, rows, gridlines=False, text_digits='B6:B%d' % (5 + len(CHEATS))), None

def build(base):
    """base: the cell style id of STYLES[0]. -> [(tab name, sheet xml, drawing xml or None)], in tab
    order after the game sheet. The host finds the checklist by its tab name."""
    return [('Cheat Sheet',) + cheat_sheet(base), ('Checklist',) + checklist(base)]
