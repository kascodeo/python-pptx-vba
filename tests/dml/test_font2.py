import pytest
from ppv.dml.fillformat import FillFormat


@pytest.fixture
def textrange2(textbox):
    return textbox.TextFrame2.TextRange


@pytest.fixture
def font2(textrange2):
    return textrange2.Font


def test_Fill(font2):
    assert isinstance(font2.Fill, FillFormat)
