from ppv.pml.presentation import Presentation


def test_parent(presentations, application):
    assert presentations.Parent is application


def test_add(presentations):
    assert isinstance(presentations.Add(), Presentation)


def test_open2007(presentations, blank_path):
    assert isinstance(presentations.Open2007(blank_path), Presentation)
