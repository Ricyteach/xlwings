# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import pytest
import inspect
from os import path, makedirs
import shutil

import xlwings as xw

@pytest.fixture
def tmpdir(scope="module"):
    tmp=path.realpath(path.join(path.dirname(inspect.getfile(inspect.currentframe())),'_tmp'))
    try:
        makedirs(tmp)
    except OSError:
        pass
    yield tmp
    try:
        shutil.rmtree(tmp)
    except OSError:
        pass

def make_app():
    return xw.App(visible=False)

@pytest.fixture(scope="module")
def app1():
    app = make_app()
    yield app
    app.kill()

@pytest.fixture(scope="module")
def app2():
    app = make_app()
    yield app
    app.kill()

def make_wb(app):
    wb = app.books.add()
    if len(wb.sheets) == 1:
        wb.sheets.add(after=1)
        wb.sheets.add(after=2)
        wb.sheets[0].select()
    return wb

@pytest.fixture
def wb1(app1):
    yield make_wb(app1)
    app1.books[-1].close()

@pytest.fixture
def wb2(app2):
    yield make_wb(app2)
    app2.books[-1].close()

def make_grp(wb, *sheets):
    grp = wb.sheets_group(*sheets)
    return grp

@pytest.fixture
def grp1(wb1):
    yield make_grp(wb1, 'sheet1', 'sheet2')

@pytest.fixture
def grp2(wb2):
    yield make_grp(wb2, 'sheet2', 'sheet3')

def test_active(grp1):
    assert grp1.active.name == grp1[0].name

def test_active_not_in_group(grp2):
    with pytest.raises(Exception):
        grp2.active

def test_index(grp1):
    assert grp1[0].name == grp1(1).name

def test_len(grp1):
    assert len(grp1) == 2

def test_del_sheet(grp1):
    name = grp1[0].name
    del grp1[0]
    assert len(grp1) == 1
    assert grp1[0].name != name

def test_iter(grp1):
    for ix, sht in enumerate(grp1):
        assert grp1[ix].name == sht.name

def test_add(grp1):
    grp1.add()
    assert len(grp1) == 3

def test_add_before(grp1):
    new_sheet = grp1.add(before='Sheet1')
    assert grp1[0].name == new_sheet.name

def test_add_after(grp1):
    grp1.add(after=len(grp1))
    assert grp1[(len(grp1) - 1)].name == grp1.active.name

    grp1.add(after=1)
    assert grp1[1].name == grp1.active.name

def test_add_default(grp1):
    current_index = grp1.active.index
    grp1.add()
    assert grp1.active.index == current_index

def test_add_named(grp1):
    grp1.add('test', before=1)
    assert grp1[0].name == 'test'

def test_add_name_already_taken(grp1):
    # does not raise exception because just adds existing sheet to group
    grp1.add('Sheet3')
    assert grp1[2].name.lower() == 'sheet3'

def test_export(grp1,tmpdir):
    filename='test.pdf'
    grp1.active['A1'].value='test'
    grp1.export(filename=path.join(tmpdir,filename), open_after_publish=False)



