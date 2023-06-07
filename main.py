# from comtypes.client import GetActiveObject
"""
import sys
import os
from contextlib import redirect_stdout

# import pymsgbox
import win32com.client
ps = win32com.client.Dispatch("PowerShape.Application")
ps.Visible = True
"""


def copy_to_points():
    # count = int(ps.evaluate('selection.number'))
    count = 10
    if count == 0:
        print('select at leas 1 object')
        sys.exit()
    # tipe = (ps.evaluate('SELECTION.TYPES'))
    tipe = '{ Solid; Point; 3dpoint; Point}'
    tipe = tipe.replace(" ", "")
    tipe = tipe.replace("{", "")
    tipe = tipe.replace("}", "")
    tipe = tipe.split(';')
    print(tipe)  # this is list of object names

    # nama = (ps.evaluate('SELECTION.NAMES'))
    nama = '{ 1; 26; 30; 99}'
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    print(nama)  # this is list of object names
    start = 0
    print('EDIT MOVE')
    print('KEEP')
    print('MOVEORIGIN')
    print('0, 0, 0,')  # start point
    for i in tipe:
        if i == 'Point':
            print(f'add {i} "{nama[start]}"')
            print('APPLY')
        start += 1
    print('DISMISS')

'''
def go_plusmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_plusmin()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
'''
if __name__ == '__main__':
    copy_to_points()

"""
---------------------------------------------------
ini adalah move copy ori nya :
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
----------------------------------------------------------------------------------
"""
