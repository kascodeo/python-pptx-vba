import pytest

from ppv.dml.font2 import Font2


@pytest.fixture
def textrange2(textbox):
    return textbox.TextFrame2.TextRange


def test_Font(textrange2):
    assert isinstance(textrange2.Font, Font2)
