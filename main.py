# from comtypes.client import GetActiveObject
import sys
import os
from contextlib import redirect_stdout
import win32com.client

ps = win32com.client.Dispatch("PowerShape.Application")
ps.Visible = True

def obj():
    count = int(ps.evaluate('selection.number'))
    if count == 0:
        print('select at leas 1 dimension')
        sys.exit()
    # print(f'jumlah dimensi yang diselect adalah {count}')

    nama = (ps.evaluate('SELECTION.NAMES'))
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    # print(nama) # this is list of dimension names
    for i in nama:
        print('select clearlist')
        print(f'add dimension "{i}"')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        # print(dim_value)
        # print(type(dim_value))

        if 0 < dim_value <= 3:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.01')
            print('TOLVALUE2 - 0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 3 < dim_value <= 6:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.012')
            print('TOLVALUE2 - 0.012')
            print('TOOLBAR DIMATTRIBUTES DISMISS')

# blm selesai yg ini

def go_plusmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{obj()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')

"""
-----------------------------------------------
#  :

-----------------------------------------------
"""
