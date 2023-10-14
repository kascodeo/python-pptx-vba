class ColorFormat():
    _srgb_pfxname = 'a:srgbClr'

    def __init__(self, fillformat):
        self._fillformat = fillformat

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
    def _textframe(self):
        return self.Parent.Parent.Parent

    @property
    def RGB(self):
        pass

    @RGB.setter
    def RGB(self, value):
        from ppv.utils.rgbcolor import MsoRGBType
        if not isinstance(value, MsoRGBType):
            raise TypeError("value must be object of MsoRGBType")

        for rpr in self._textframe.get_rpr():
            fill = self.Parent.get_fill_e(rpr)
            if fill is not None:
                srgb = fill.findqn(self._srgb_pfxname)
                if srgb is None:
                    srgb = fill.makeelement(fill.qn(self._srgb_pfxname))
                    fill.insert(0, srgb)
                srgb.set('val', value.hexstr.upper())
