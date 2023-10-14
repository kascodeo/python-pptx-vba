from .shapes import Shapes
from ..dml.cell import Cell


class Shape():
    def __init__(self, parent, e):
        self._parent = parent
        self._e = e
        self._textframe = None
        self._textframe2 = None
        self._table = None

    @property
    def e(self):
        return self._e

    @property
    def Parent(self):
        if isinstance(self._parent, Shapes):
            return self._parent.Parent
        if isinstance(self._parent, Cell):
            return self._parent.Parent.Parent.Parent

    @property
    def HasTextFrame(self):
        if self.e.tag == self.e.qn('p:sp'):
            pfxtag = 'p:txBody'
        elif self.e.tag == self.e.qn('a:tc'):
            pfxtag = 'a:txBody'
        return self.e.findqn(pfxtag) is not None

    @property
    def TextFrame(self):
        from ppv.pml.textframe import TextFrame
        if self.HasTextFrame:
            if self._textframe is None:
                if self.e.tag == self.e.qn('p:sp'):
                    pfxtag = 'p:txBody'
                elif self.e.tag == self.e.qn('a:tc'):
                    pfxtag = 'a:txBody'
                self._textframe = TextFrame(self, self.e.findqn(pfxtag))
            return self._textframe

    @property
    def TextFrame2(self):
        from ppv.pml.textframe2 import TextFrame2

        if self._textframe2 is None:
            self._textframe2 = TextFrame2(self, self.e.findqn('p:txBody'))
        return self._textframe2

    @property
    def HasChart(self):
        pass

    @property
    def HasTable(self):
        if self.e.find('.//'+self.e.qn('a:tbl')) is not None:
            return True
        return False

    @property
    def Table(self):
        from ..dml.table import Table
        if self.HasTable:
            self._table = Table(self)
            return self._table

    @property
    def Height(self):
        pass

    @property
    def Adjustments(self):
        pass

    @property
    def AutoShapeType(self):
        pass

    @property
    def Id(self):
        pass

    @property
    def Left(self):
        pass

    @property
    def Name(self):
        pass

    @property
    def PlaceholderFormat(self):
        pass

    @property
    def Rotation(self):
        pass

    @property
    def Shadow(self):
        pass

    @property
    def Top(self):
        pass

    @property
    def Type(self):
        pass

    @property
    def Width(self):
        pass

    @property
    def Fill(self):
        pass

    @property
    def Line(self):
        pass

    @property
    def OLEFormat(self):
        pass
