
class TestPpv():
    def test_ppv_has_application(self,):
        from ppv import Application
        from ppv.pml.application import Application as Application_
        assert Application is Application_()

    def test_ppv_has_presentations(self):
        from ppv import Presentations
        from ppv import Application
        assert Presentations is Application.Presentations
