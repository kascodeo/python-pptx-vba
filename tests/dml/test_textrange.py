import pytest

from ppv.dml.font import Font
from ppv.dml.textrange import TextRange


@pytest.fixture
def textrange(textbox):
    return textbox.TextFrame.TextRange


def test_Text(textrange):
    assert textrange.Text == 'Text'


def test_Parent(textrange, textbox):
    assert textrange.Parent is textbox.TextFrame


def test_e(textrange):
    assert textrange.e.ln == 'txBody'


def test_Font(textrange):
    assert isinstance(textrange.Font, Font)


def test_Text_set(textrange):
    newtext = "Some New Text"
    textrange.Text = newtext
    assert textrange.Text == newtext


def test_Characters(textrange):
    assert isinstance(textrange.Characters(), TextRange)


def test_Characters_Text(textrange):
    assert textrange.Text == textrange.Characters().Text


def test_InsertAfter_return(textrange):
    assert isinstance(textrange.InsertAfter("-After-"), TextRange)


def test_InsertAfter_return_Text(textrange):
    assert textrange.InsertAfter("-After-").Text == '-After-'
