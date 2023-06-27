# from comtypes.client import GetActiveObject
import sys
import os
from contextlib import redirect_stdout
import win32com.client
ps = win32com.client.Dispatch("PowerShape.Application")
ps.Visible = True


def copy_to_points():
    ps.exec('FILTERBUTTON FILTERITEMS')
    ps.exec('SelectType  Point')
    ps.exec('ALL')
    ps.exec('ACCEPT')
    ps.exec('EVERYTHING PARTIALBOX')
    count = int(ps.evaluate('selection.number'))
    # count = 10
    if count == 0:
        print('select at leas 1 object')
        sys.exit()
    tipe = (ps.evaluate('SELECTION.TYPES'))
    # tipe = '{ Solid; Point; 3dpoint; Point}'
    tipe = tipe.replace(" ", "")
    tipe = tipe.replace("{", "")
    tipe = tipe.replace("}", "")
    tipe = tipe.split(';')
    # print(tipe)  # this is list of object names

    nama = (ps.evaluate('SELECTION.NAMES'))
    # nama = '{ 1; 26; 30; 99}'
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    # print(nama)  # this is list of object names
    start = 0
    print('RESTORE_SELECTION')
    ListPoint = []
    for i in tipe:
        if i == 'Point':
            xyz = str(ps.evaluate(f'point[{nama[start]}].position'))
            xyz = xyz.replace(',', '')
            xyz = xyz.replace('(', '')
            xyz = xyz.replace(')', '')
            xyz = xyz.replace('.', ',')
            ListPoint.append(xyz)
        start += 1

    for i in ListPoint:
        print('EDIT MOVE')
        print('KEEP')
        print('MOVEORIGIN')
        print('0, 0, 0,')
        print(i)
        print('APPLY')
        print('DISMISS')
        print('RESTORE_SELECTION')
        print('RESTORE_SELECTION')

        start += 1
    # print(ListPoint)
    return

def go():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{copy_to_points()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')


if __name__ == '__main__':
    copy_to_points()
    go()

'''
def go_plusmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_plusmin()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
'''

"""
----------algoritma yang dipake----------
$$ add Surface "2"
$$ add Solid "1"
$$ add Arc "4"
$$ add Workplane "1"
$$ add Line "8"
$$ add CompCurve "1"
$$ add Component "NEW_MODEL_1_assembly" "2_1"
FILTERBUTTON FILTERITEMS
SelectType  Point
ALL
ACCEPT
EVERYTHING PARTIALBOX
RESTORE_SELECTION
EDIT MOVE
KEEP
MOVEORIGIN
-0, -0, 0,
$$ 11,7 7,5 0,
APPLY
$$ 9, 10, 0,
APPLY
$$ 27, 9, 0,
APPLY
$$ 39, 0, 0,
APPLY
DISMISS
$$ select clearlist

----------format output pointnya----------
LET $koor = point ['2'].position 
PRINT $koor
output : 
[20,700000, 17,500000, 0,000000]

----------move copy ori nya----------
add Solid "1"
EDIT MOVE
KEEP
MOVEORIGIN
0, 0, 0,
-50, -0, 0,
APPLY
-46, 1, 0,
APPLY
4, 42, 0,
APPLY
43, -1, 0,
APPLY
DISMISS
"""
