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
def copy_to_point():
    # count = int(ps.evaluate('selection.number'))
    count = 10
    if count == 0:
        print('select at leas 1 dimension')
        sys.exit()
    print(f'jumlah object yang diselect adalah {count}')

    # nama = (ps.evaluate('SELECTION.NAMES'))
    nama = '{ Solid; Point }'
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    print(nama) # this is list of object names
    '''
    for i in nama:
        print('select clearlist')
        print(f'add ---disinitipeobjecnya--- "{i}"')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        # print(dim_value)
        # print(type(dim_value))
    '''
'''
def go_plusmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_plusmin()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
'''
