class FillFormat():
    _solid = 'solidFill'
    _solid_pfxname = 'a:solidFill'

    def __init__(self, font) -> None:
        from .colorformat import ColorFormat
        self._font = font
        self._forecolor = ColorFormat(self)

    def Background(self):
        pass

    def Patterned(self):
        pass

    def Solid(self):
        textrange = self.Parent.Parent
        textrange.isolate()
        for rpr in textrange.get_rpr_in_range():
            fill = self.get_fill_e(rpr)
            if fill is not None and fill.ln != self._solid:
                fill.getparent().remove(fill)
            fill = rpr.findqn(self._solid_pfxname)
            if fill is None:
                fill = rpr.makeelement(rpr.qn(self._solid_pfxname))
                rpr.insert(0, fill)

    def get_fill_e(self, rpr):
        for e in rpr:
            if 'Fill' in e.ln:
                return e

    @property
    def BackColor(self):
        pass

    @property
    def GradientAngle(self):
        pass

    @property
    def GradientColorType(self):
        pass

    @property
    def GradientDegree(self):
        pass

    @property
    def GradientStops(self):
        pass

    @property
    def GradientStyle(self):
        pass

    @property
    def Pattern(self):
        pass

    @property
    def ForeColor(self):
        return self._forecolor

    @property
    def Parent(self):
        return self._font
