def test_Parent(blank_pres):
    assert blank_pres.Slides.Parent is blank_pres


def test_Application(blank_pres, application):
    assert blank_pres.Slides.Application is application


def test_Count(blank_pres):
    assert blank_pres.Slides.Count == 1


def test_AddSlide(blank_pres):
    before = blank_pres.Slides.Count
    blank_pres.Slides.AddSlide(2, blank_pres.SlideMaster.CustomLayouts.Item(5))
    assert blank_pres.Slides.Count == before + 1


def test_Item(blank_pres):
    from ppv.pml.slide import Slide
    assert isinstance(blank_pres.Slides.Item(1), Slide)


def test_FindBySlideID(blank_pres):
    assert blank_pres.Slides.Item(1) is blank_pres.Slides.FindBySlideID(256)


def test_get_next_uri_str(blank_pres):
    assert blank_pres.Slides.get_next_uri_str() == "/ppt/slides/slide2.xml"
