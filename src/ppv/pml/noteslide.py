from ..core.xmltypeobj import XmlTypeobj


class NoteSlide(XmlTypeobj):
    type = ("application/vnd.openxmlformats-officedocument.presentationml."
            "notesSlide+xml")

    reltype = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/notesSlide")

    def __init__(self, part):
        super().__init__(part)
