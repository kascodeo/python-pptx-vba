import pytest
from ppv.core.package import Package
from ppv.pml.presentation import Presentation
from ppv.pml.handoutmaster import HandoutMaster
from ppv.pml.notesmaster import NotesMaster
from ppv.pml.slide import Slide
from ppv.pml.slidemaster import SlideMaster
from ppv.pml.presentationpr import PresentationPr
from ppv.pml.viewpr import ViewPr
from ppv.pml.tablestyles import TableStyles
from ppv.pml.slidelayout import SlideLayout
from ppv.od.theme import Theme
from ppv.dml.chartspace import ChartSpace
from ppv.mso.chartstyle import ChartStyle
from ppv.mso.colorstyle import ColorStyle
from ppv.od.properties import Properties
from ppv.img.image import Image


@pytest.mark.parametrize("class_", [Presentation,
                                    HandoutMaster,
                                    NotesMaster,
                                    Slide,
                                    SlideMaster,
                                    PresentationPr,
                                    ViewPr,
                                    TableStyles,
                                    SlideLayout,
                                    Theme,
                                    ChartSpace,
                                    ChartStyle,
                                    ColorStyle,
                                    Properties,
                                    ])
def test_hook_is_registered(blank_path, class_):
    package = Package(blank_path)
    assert package._part_hooks.get(class_.type) is class_


def test_hook_is_registered_image(blank_path):
    package = Package(blank_path)
    for type in Image.types:
        assert package._part_hooks.get(type) is Image


def test_read(blank_path):
    package = Package(blank_path)
    assert package.document is None
    package.read()
    assert isinstance(package.document, Presentation)
