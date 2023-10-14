from ..core.xmltypeobj import XmlTypeobj
from ..core.package import Package


class Presentation(XmlTypeobj):
    uri_str = "/ppt/presentation.xml"
    type = ("application/vnd.openxmlformats-officedocument.presentationml."
            "presentation.main+xml")

    @classmethod
    def create(cls, path):
        pkg = Package(path).read()
        self = pkg.document
        self._pkg = pkg
        self._slides = None
        return self

    @property
    def Parent(self):
        return self.Application.Presentations

    def SaveAs(self, FileName):
        self._pkg.write(FileName)

    @property
    def BuiltInDocumentProperties(self):
        return self._pkg.core_properties

    @property
    def HasHandoutMaster(self):
        return self.get_rid('p:handoutMasterIdLst',
                            'p:handoutMasterId') is not None

    @property
    def HandoutMaster(self):
        return self.get_related_typeobj(self.get_rid('p:handoutMasterIdLst',
                                                     'p:handoutMasterId'))

    @property
    def HasNotesMaster(self):
        return self.get_rid('p:notesMasterIdLst',
                            'p:notesMasterId') is not None

    @property
    def NotesMaster(self):
        return self.get_related_typeobj(self.get_rid('p:notesMasterIdLst',
                                                     'p:notesMasterId'))

    @property
    def SlideMaster(self):
        return self.get_related_typeobj(self.get_rid('p:sldMasterIdLst',
                                                     'p:sldMasterId'))

    @property
    def Slides(self):
        from ppv.pml.slides import Slides
        if self._slides is None:
            self._slides = Slides(self)
        return self._slides

    def get_rid(self, pfxname_lst, pfxname_id):
        lst_e = self.e.findqn(pfxname_lst)
        if lst_e is not None:
            id_e = lst_e.findqn(pfxname_id)
            if id_e is not None:
                return id_e.getqn('r:id')

    @property
    def SlideMasters(self):
        pass
