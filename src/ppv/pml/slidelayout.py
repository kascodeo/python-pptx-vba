from ..core.xmltypeobj import XmlTypeobj


class SlideLayout(XmlTypeobj):
    type = ("application/vnd.openxmlformats-officedocument.presentationml."
            "slideLayout+xml")
    reltype = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/slideLayout")

    def Delete(self):
        pass

    @property
    def Parent(self):
        pass

    @property
    def FollowMasterBackground(self):
        pass

    @property
    def Shapes(self):
        from ppv.pml.shapes import Shapes
        if not hasattr(self, "_shapes"):
            self._shapes = Shapes(self)
        return self._shapes
