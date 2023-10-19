import pytest
from ppv.dml.textrange import TextRange
from ppv.enum import MsoTriState
from ppv.pml.shape import Shape


@pytest.fixture
def textframe(textbox):
    return textbox.TextFrame


def test_Parent(textframe):
    assert isinstance(textframe.Parent, Shape)


def test_e(textframe):
    assert textframe.e.ln == 'txBody'


def test_HasText(textframe):
    assert textframe.HasText is True


def test_TextRange(textframe):
    assert isinstance(textframe.TextRange, TextRange)


def test_WordWrap(textframe):
    assert textframe.WordWrap == MsoTriState.msoFalse


def test_WordWrap_set(textframe):
    textframe.WordWrap = MsoTriState.msoTrue
    assert textframe.WordWrap == MsoTriState.msoTrue
