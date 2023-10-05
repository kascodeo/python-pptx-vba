def test_Parent(slide):
    assert slide.Shapes.Parent is slide


def test_Application(slide, application):
    assert slide.Shapes.Application is application


def test_Count(shapes_pres):
    assert shapes_pres.Slides.Item(1).Shapes.Count == 13
