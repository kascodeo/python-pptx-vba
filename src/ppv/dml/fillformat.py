class FillFormat():
    _solidFill = 'solidFill'
    _a_solidFill = 'a:solidFill'
    _a_noFill = 'a:noFill'

    def __init__(self, parent, e=None) -> None:
        from .colorformat import ColorFormat
        self._parent = parent
        self._forecolor = ColorFormat(self, 'forecolor')
        self._backcolor = ColorFormat(self, 'backcolor')

    def Background(self):
        pass

    def Patterned(self):
        pass

    def Solid(self):
        textrange = self.Parent.Parent
        textrange.isolate()
        for rpr in textrange.get_rpr_in_range():
            fill = self.get_fill_e(rpr)
            if fill is not None and fill.ln != self._solidFill:
                fill.getparent().remove(fill)
            fill = rpr.findqn(self._a_solidFill)
            if fill is None:
                fill = rpr.makeelement(rpr.qn(self._a_solidFill))
                rpr.insert(0, fill)

    def get_fill_e(self, rpr):
        for e in rpr:
            if 'Fill' in e.ln:
                return e

    @property
    def BackColor(self):
        return self._backcolor

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
        return self._parent

    def get_or_create_solidFill(self, e, index=-1):
        solidFill = e.findqn(self._a_solidFill)
        if solidFill is None:
            solidFill = e.makeelement(e.qn(self._a_solidFill))
            e.insert(index, solidFill)
        return solidFill

    def set_solidFill(self, e):
        index = -1
        i = -1
        for e_ in e:
            i += 1
            if 'Fill' in e_.ln and e_.ln != self._solidFill:
                e_.getparent().remove(e_)
                index = i
        solidFill = self.get_or_create_solidFill(e, index)
        return solidFill

    def set_noFill(self, e):
        index = -1
        for i, e_ in enumerate(e.getchildren()):
            if 'Fill' in e_.ln:
                e_.getparent().remove(e_)
                index = i
        e.insert(index, e.makeelement(e.qn(self._a_noFill)))
