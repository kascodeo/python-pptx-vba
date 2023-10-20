import pytest

from ppv.dml.row import Row


@pytest.fixture
def rows(table):
    return table.Rows


def test_e(rows, table):
    assert rows.e is table.e


def test_e_lst_tr(rows):
    assert isinstance(rows.e_lst_tr, list)
    assert rows.e_lst_tr[0].ln == 'tr'


def test_get_all(rows):
    assert isinstance(rows.get_all(), list)


def test_Add(rows):
    row = rows.Add()
    assert row.e is rows.e_lst_tr[-1]


def test_Count(rows):
    assert rows.Count == 2


def test_Item(rows):
    assert isinstance(rows.Item(1), Row)


def test_Item_e(rows):
    assert rows.Item(1).e is rows.e_lst_tr[0]
