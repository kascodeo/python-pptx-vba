import pytest

from ppv.dml.cell import Cell
from ppv.dml.columns import Columns
from ppv.dml.rows import Rows
from ppv.pml.shape import Shape


@pytest.fixture
def table(shape_of_table):
    return shape_of_table.Table


def test_e(table):
    assert table.e.ln == 'tbl'


def test_e_tbl(table):
    assert table.e_tbl.ln == 'tbl'


def test_e_tblPr(table):
    assert table.e_tblPr.ln == 'tblPr'


def test_Cell(table):
    assert isinstance(table.Cell(1, 1), Cell)


def test_Columns(table):
    assert isinstance(table.Columns, Columns)


def test_FirstCol(table):
    assert table.FirstCol is False


def test_FirstCol_set(table):
    table.FirstCol = True
    assert table.FirstCol is True
    table.FirstCol = False
    assert table.FirstCol is False


def test_FirstRow(table):
    assert table.FirstRow is True


def test_FirstRow_set(table):
    table.FirstRow = False
    assert table.FirstRow is False
    table.FirstRow = True
    assert table.FirstRow is True


def test_HorizBanding(table):
    assert table.HorizBanding is True


def test_HorizBanding_set(table):
    table.HorizBanding = False
    assert table.HorizBanding is False
    table.HorizBanding = True
    assert table.HorizBanding is True


def test_LastCol(table):
    assert table.LastCol is False


def test_LastCol_set(table):
    table.LastCol = True
    assert table.LastCol is True
    table.LastCol = False
    assert table.LastCol is False


def test_LastRow(table):
    assert table.LastRow is False


def test_LastRow_set(table):
    table.LastRow = True
    assert table.LastRow is True
    table.LastRow = False
    assert table.LastRow is False


def test_Parent(table):
    assert isinstance(table.Parent, Shape)


def test_Rows(table):
    assert isinstance(table.Rows, Rows)


def test_VertBanding(table):
    assert table.VertBanding is False


def test_VertBanding_set(table):
    table.VertBanding = True
    assert table.VertBanding is True
    table.VertBanding = False
    assert table.VertBanding is False


def test_get_tblPr_attrib(table):
    attr = 'bandRow'
    assert table.get_tblPr_attrib(attr) is True


def test_set_tblPr_attrib(table):
    attr = 'bandRow'
    table.set_tblPr_attrib(attr, False)
    assert table.get_tblPr_attrib(attr) is False
