def test_CustomLayouts(blank_pres):
    from ppv.pml.customlayouts import CustomLayouts
    assert isinstance(blank_pres.SlideMaster.CustomLayouts, CustomLayouts)
