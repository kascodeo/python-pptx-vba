def test_Count(blank_pres):
    assert blank_pres.SlideMaster.CustomLayouts.Count == 11


def test_Parent(blank_pres):
    assert blank_pres.SlideMaster.CustomLayouts.Parent is blank_pres.\
        SlideMaster


def test_Application(blank_pres, application):
    assert blank_pres.SlideMaster.CustomLayouts.Application is application
