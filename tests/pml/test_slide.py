def test_Parent(blank_pres, slide):
    assert slide.Parent is blank_pres


def test_Application(slide, application):
    assert slide.Application is application


def test_CustomLayout(slide, slide_layout):
    assert slide.CustomLayout is slide_layout


def test_Shapes(slide):
    from ppv.pml.shapes import Shapes
    assert isinstance(slide.Shapes, Shapes)


def test_SlideID(slide):
    assert slide.SlideID == 256


def test_SlideIndex(slide):
    assert slide.SlideIndex == 1
