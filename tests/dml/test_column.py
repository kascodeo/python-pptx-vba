import pytest
from ppv.dml.cellrange import CellRange


@pytest.fixture
def columns(table):
    return table.Columns


@pytest.fixture
def column(columns):
    return columns.Item(1)


def test_e(column):
    assert column.e.ln == 'gridCol'


def test_Parent(column, table):
    assert column.Parent is table


def test_Cells(column):
    assert isinstance(column.Cells, CellRange)


def test_index(column):
    assert column.index == 0


def test_rowId(column):
    assert column.colId == '4120647543'
