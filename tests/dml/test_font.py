import pytest


from ppv.dml.textrange import TextRange
from ppv.enum.MsoTriState import MsoTriState


@pytest.fixture
def textrange(textbox):
    return textbox.TextFrame.TextRange


@pytest.fixture
def font(textrange):
    return textrange.Font


def test_Parent(font):
    assert isinstance(font.Parent, TextRange)


def test_Bold_return(font):
    assert isinstance(font.Bold, MsoTriState)


def test_Bold_return_msoFalse(font):
    assert font.Bold is MsoTriState.msoFalse


def test_Bold_return_msoTrue(font):
    font.Bold = MsoTriState.msoTrue
    assert font.Bold is MsoTriState.msoTrue


def test_Italic_return(font):
    assert isinstance(font.Italic, MsoTriState)


def test_Italic_return_msoFalse(font):
    assert font.Italic is MsoTriState.msoFalse


def test_Italic_return_msoTrue(font):
    font.Italic = MsoTriState.msoTrue
    assert font.Italic is MsoTriState.msoTrue


def test_Name_return(font):
    assert isinstance(font.Name, str)
    assert font.Name == font.default_name


def test_Name_return_after_change(font):
    name = "Arial"
    font.Name = name
    assert font.Name == name


def test_Size_return(font):
    assert font.Size == '18'


def test_Size_return_after_change(font):
    font.Size = 20.0
    assert font.Size == '20'
