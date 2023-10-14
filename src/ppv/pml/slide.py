from ..core.xmltypeobj import XmlTypeobj


class Slide(XmlTypeobj):
    type = ("application/vnd.openxmlformats-officedocument.presentationml."
            "slide+xml")

    reltype = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/slide")

    @property
    def Background(self):
        pass

    @property
    def HasNotesPage(self):
        pass

    @property
    def Name(self):
        pass

    @property
    def NotesPage(self):
        pass

    @property
    def CustomLayout(self):
        from ppv.pml.slidelayout import SlideLayout
        part = self.part.get_related_parts_by_reltype(SlideLayout.reltype)[0]
        return part.typeobj

    @property
    def Shapes(self):
        from ppv.pml.shapes import Shapes
        if not hasattr(self, "_shapes"):
            self._shapes = Shapes(self)
        return self._shapes

    @property
    def SlideID(self):
        return self._get_id_index("id")

    @property
    def SlideIndex(self):
        return self._get_id_index("index")

    def _get_id_index(self, which="id"):
        for i, sldId in enumerate(self.Parent.Slides.e):
            rid = sldId.getqn('r:id')
            slide = self.Parent.get_related_typeobj(rid)
            if slide is self:
                if which == "id":
                    return int(sldId.get('id'))
                elif which == "index":
                    return i + 1
