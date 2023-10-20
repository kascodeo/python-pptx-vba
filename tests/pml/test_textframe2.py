
import pytest
from ppv.dml.textrange2 import TextRange2


@pytest.fixture
def textframe2(textbox):
    return textbox.TextFrame2


def test_TextRange(textframe2):
    assert isinstance(textframe2.TextRange, TextRange2)
