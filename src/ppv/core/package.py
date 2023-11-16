from opc import Package as OPCPackage


class Package(OPCPackage):
    def __init__(self, *args):
        from ..pml.presentation import Presentation
        from ..pml.handoutmaster import HandoutMaster
        from ..pml.notesmaster import NotesMaster
        from ..pml.slide import Slide
        from ..pml.noteslide import NoteSlide
        from ..pml.master import Master
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
        self.register_part_hook(NoteSlide.type, NoteSlide)
        self.register_part_hook(Master.type, Master)
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

    def get_uri_type_picture(self, filename):
        ext = filename.split('.')[-1].lower()
        i = 0
        for uri_str in self._parts:
            pattern = '/ppt/media/image'
            if uri_str.startswith(pattern):
                i_ = int(uri_str.split('.')[0].replace(pattern, ''))
                if i_ > i:
                    i = i_
        return '{}{}.{}'.format(pattern, i+1, ext), 'image/'+ext

    def get_available_uri(self, base_uri):
        base, ext = base_uri.split('.')
        ext = ext.lower()
        i = 0
        for uri_str in self._parts:
            if uri_str.startswith(base):
                try:
                    i_ = int(uri_str.split('.')[0].replace(base, ''))
                    if i_ > i:
                        i = i_
                except ValueError:
                    pass
        return base + str(i+1)+'.'+ext

    def add_part_picture(self, filename):
        uri_str, type_ = self.get_uri_type_picture(filename)
        part = self.add_part(uri_str, type_)
        return part
