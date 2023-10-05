
from ppv.pml.application import Application


def test_has_presentations():
    from ppv import Application as Application_
    assert isinstance(Application_, Application)


def test_singleton():
    assert Application() is Application()


def test_presentations_is_identical_for_application():
    assert Application().Presentations is Application().Presentations
