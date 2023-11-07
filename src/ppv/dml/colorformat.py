from ppv.dml.font2 import Font2
from ppv.pml.shape import Shape
from ppv.utils.rgbcolor import MsoRGBType


class ColorFormat():
    _srgb_pfxname = 'a:srgbClr'
    _kind_forecolor = 'forecolor'
    _kind_backcolor = 'backcolor'
    a_srgbClr = 'a:srgbClr'

    def __init__(self, fillformat, kind):
        self._fillformat = fillformat
        self._kind = kind

    @property
    def Brightness(self):
        return self._fillformat

    @property
    def ObjectThemeColor(self):
        return self._fillformat

    @property
    def Type(self):
        return self._fillformat

    @property
    def Parent(self):
        return self._fillformat

    @property
    def _textrange(self):
        return self.Parent.Parent.Parent

    @property
    def _shape(self):
        return self.Parent.Parent

    @property
    def RGB(self):
        pass

    @RGB.setter
    def RGB(self, value):
        if not (isinstance(value, MsoRGBType) or value is None):
            raise TypeError("value must be object of MsoRGBType")

        if isinstance(self.Parent.Parent, Font2):
            self.change_font_color(value)

        if isinstance(self.Parent.Parent, Shape):
            self.change_shape_color(value)

    def change_font_color(self, value):
        if self._kind == 'forecolor':
            self._textrange.update_lst_rpr(
                'rgb', self._get_none_or_rgbhex(value))

    def change_shape_color(self, value):
        sp = self.Parent.Parent
        # spPr = sp.spPr
        spPr = sp.e_Pr
        value = self._get_none_or_rgbhex(value)
        if self._kind == self._kind_backcolor:
            if value is None:
                self.Parent.set_noFill(spPr)
            else:
                solidFill = self.Parent.set_solidFill(spPr)
                self.set_srgbClr(solidFill, value)

    def set_srgbClr(self, e, value):
        srgbClr = e.findqn(self.a_srgbClr)
        if srgbClr is None:
            srgbClr = e.makeelement(e.qn(self.a_srgbClr))
            e.insert(0, srgbClr)
        srgbClr.set('val', value)

    def _get_none_or_rgbhex(self, value):
        return None if value is None else value.hexstr.upper()
