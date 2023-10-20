import pytest

from ppv.dml.column import Column


@pytest.fixture
def columns(table):
    return table.Columns


def test_e(columns, table):
    assert columns.e is table.e


def test_e_tblGrid(columns):
    assert columns.e_tblGrid.ln == 'tblGrid'


def test_e_lst_gridCol(columns):
    assert isinstance(columns.e_lst_gridCol, list)
    assert columns.e_lst_gridCol[0].ln == 'gridCol'


def test_get_all(columns):
    assert isinstance(columns.get_all(), list)


def test_Add(columns):
    column = columns.Add()
    assert column.e is columns.e_lst_gridCol[-1]


def test_Count(columns):
    assert columns.Count == 2


def test_Item(columns):
    assert isinstance(columns.Item(1), Column)


def test_Item_e(columns):
    assert columns.Item(1).e is columns.e_lst_gridCol[0]
