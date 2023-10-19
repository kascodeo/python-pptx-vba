
from ppv.pml.textframe import TextFrame
from ppv.pml.textframe2 import TextFrame2
from ppv.pml.slide import Slide


def test_Parent(textbox):
    assert isinstance(textbox.Parent, Slide)


def test_HasTextFrame(textbox):
    assert textbox.HasTextFrame is True


def test_TextFrame(textbox):
    assert isinstance(textbox.TextFrame, TextFrame)


def test_TextFrame2(textbox):
    assert isinstance(textbox.TextFrame2, TextFrame2)


def test_HasTable_false(textbox):
    assert textbox.HasTable is False
