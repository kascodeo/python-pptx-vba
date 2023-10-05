from opc import Package as OPCPackage


class Package(OPCPackage):
    def __init__(self, *args):
        from ..pml.presentation import Presentation
        from ..pml.handoutmaster import HandoutMaster
        from ..pml.notesmaster import NotesMaster
        from ..pml.slide import Slide
        from ..pml.slidemaster import SlideMaster
        from ..pml.presentationpr import PresentationPr
        from ..pml.viewpr import ViewPr
        from ..pml.tablestyles import TableStyles
        from ..pml.slidelayout import SlideLayout
        from ..od.theme import Theme
        from ..dml.chartspace import ChartSpace
        from ..mso.chartstyle import ChartStyle
        from ..mso.colorstyle import ColorStyle
        from ..od.properties import Properties
        from ..img.image import Image

        super().__init__(*args)
        self.register_part_hook(Presentation.type, Presentation)
        self.register_part_hook(HandoutMaster.type, HandoutMaster)
        self.register_part_hook(NotesMaster.type, NotesMaster)
        self.register_part_hook(Slide.type, Slide)
        self.register_part_hook(SlideMaster.type, SlideMaster)
        self.register_part_hook(SlideLayout.type, SlideLayout)
        self.register_part_hook(PresentationPr.type, PresentationPr)
        self.register_part_hook(ViewPr.type, ViewPr)
        self.register_part_hook(TableStyles.type, TableStyles)
        self.register_part_hook(Theme.type, Theme)
        self.register_part_hook(ChartSpace.type, ChartSpace)
        self.register_part_hook(ChartStyle.type, ChartStyle)
        self.register_part_hook(ColorStyle.type, ColorStyle)
        self.register_part_hook(Properties.type, Properties)
        for type in Image.types:
            self.register_part_hook(type, Image)

        self._document = None

    def read(self, *args, **kwargs):
        super().read(*args, **kwargs)
        self._document = self.get_part('/ppt/presentation.xml').typeobj
        return self

    @property
    def document(self):
        return self._document
