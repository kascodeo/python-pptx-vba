from ..core.xmltypeobj import XmlTypeobj
from ppv.pml.noteslide import NoteSlide


class Slide(XmlTypeobj):
    type = ("application/vnd.openxmlformats-officedocument.presentationml."
            "slide+xml")

    reltype = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/slide")

    def __init__(self, part):
        super().__init__(part)
        self._background = None

    @property
    def Background(self):
        from .shaperange import ShapeRange
        if self._background is None:
            self._background = ShapeRange(self, self.e)

    @property
    def HasNotesPage(self):
        parts = self.part.get_related_parts_by_reltype(NoteSlide.reltype)
        return True if len(parts) == 1 else False

    @property
    def Name(self):
        return 'slide'+str(self.SlideIndex)

    @property
    def NotesPage(self):
        if self.HasNotesPage:
            part = self.part.get_related_parts_by_reltype(NoteSlide.reltype)[0]
            return part.typeobj

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
