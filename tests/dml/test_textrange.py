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


def test_InsertAfter_full_text(textrange):
    textrange.InsertAfter("-After-")
    assert textrange.Text == 'Text-After-'


def test_InsertBefore_return(textrange):
    assert isinstance(textrange.InsertBefore("-Before-"), TextRange)


def test_InsertBefore_return_Text(textrange):
    assert textrange.InsertBefore("-Before-").Text == '-Before-'


def test_InsertBefore_full_text(textrange):
    tr = textrange.InsertAfter("-After-")
    tr.InsertBefore("-Before-")
    assert textrange.Text == "Text-Before--After-"


def test_Text_from_empty_text(textrange):
    p, r, t = textrange.get_or_create_p_r_t()
    r.getparent().remove(r)
    txt = "From empty text"
    textrange.Text = txt
    assert textrange.Text == txt


def test_InsertAfter_from_empty_text(textrange):
    p, r, t = textrange.get_or_create_p_r_t()
    r.getparent().remove(r)
    txt = "From empty text"
    textrange.InsertAfter(txt)
    assert textrange.Text == txt


def test_InsertBefore_from_empty_text(textrange):
    p, r, t = textrange.get_or_create_p_r_t()
    r.getparent().remove(r)
    txt = "From empty text"
    textrange.InsertBefore(txt)
    assert textrange.Text == txt
