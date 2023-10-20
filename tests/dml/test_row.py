import pytest
from ppv.dml.cellrange import CellRange


@pytest.fixture
def rows(table):
    return table.Rows


@pytest.fixture
def row(rows):
    return rows.Item(1)


def test_e(row):
    assert row.e.ln == 'tr'


def test_Parent(row, table):
    assert row.Parent is table


def test_Cells(row):
    assert isinstance(row.Cells, CellRange)


def test_index(row):
    assert row.index == 0


def test_rowId(row):
    assert row.rowId == '607152856'
