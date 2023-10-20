import pytest

from ppv.pml.shape import Shape


@pytest.fixture
def table(shape_of_table):
    return shape_of_table.Table


@pytest.fixture
def cell(table):
    return table.Cell(1, 1)


def test_e(cell):
    assert cell.e.ln == 'tc'


def test_Parent(cell, table):
    assert cell.Parent is table


def test_Shape(cell):
    assert isinstance(cell.Shape, Shape)
