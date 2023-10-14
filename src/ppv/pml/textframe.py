from ppv.enum.MsoTriState import MsoTriState


class TextFrame():
    def __init__(self, shape, e):
        self._shape = shape
        self._e = e

    def DeleteText(self):
        pass

    @property
    def AutoSize(self):
        pass

    @property
    def Parent(self):
        return self._shape

    @property
    def e(self):
        return self._e

    @property
    def HasText(self):
        return self.e.findqn('a:p') is not None

    @property
    def TextRange(self):
        from ppv.dml.textrange import TextRange
        return TextRange(self)

    @property
    def WordWrap(self):
        if self.e.get('wrap', 'none') == 'none':
            return MsoTriState.msoFalse
        return MsoTriState.msoTrue

    @WordWrap.setter
    def WordWrap(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value]:
            raise ValueError("val must be msoTrue|msoFalse")
        if val in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            val = 'square'
        if val in [MsoTriState.msoFalse, MsoTriState.msoFalse.value]:
            val = 'none'
        self.e.set('wrap', val)

    @property
    def MarginBottom(self):
        pass

    @property
    def MarginLeft(self):
        pass

    @property
    def MarginRight(self):
        pass

    @property
    def MarginTop(self):
        pass

    @property
    def VerticalAnchor(self):
        pass
