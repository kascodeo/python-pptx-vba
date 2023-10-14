class Table():
    bandCol = 'bandCol'
    bandRow = 'bandRow'
    firstRow = 'firstRow'
    firstCol = 'firstCol'
    lastRow = 'lastRow'
    lastCol = 'lastCol'

    def __init__(self, shape):
        self._shape = shape
        self._columns = None
        self._rows = None

    @property
    def e(self):
        return self.e_tbl

    @property
    def e_tbl(self):
        e = self.Parent.e
        return e.find('.//'+e.qn('a:tbl'))

    @property
    def e_tblPr(self):
        return self.e_tbl.findqn('a:tblPr')

    def Cell(self, Row, Column):
        return self.Rows.Item(Row).Cells.Item(Column)

    @property
    def Columns(self):
        from .columns import Columns
        if self._columns is None:
            self._columns = Columns(self)
        return self._columns

    @property
    def FirstCol(self):
        return self.get_tblPr_attrib(self.firstCol)

    @FirstCol.setter
    def FirstCol(self, value):
        self.set_tblPr_attrib(self.firstCol, value)

    @property
    def FirstRow(self):
        return self.get_tblPr_attrib(self.firstRow)

    @FirstRow.setter
    def FirstRow(self, value):
        self.set_tblPr_attrib(self.firstRow, value)

    @property
    def HorizBanding(self):
        return self.get_tblPr_attrib(self.bandRow)

    @HorizBanding.setter
    def HorizBanding(self, value):
        self.set_tblPr_attrib(self.bandRow, value)

    @property
    def LastCol(self):
        return self.get_tblPr_attrib(self.lastCol)

    @LastCol.setter
    def LastCol(self, value):
        self.set_tblPr_attrib(self.lastCol, value)

    @property
    def LastRow(self):
        return self.get_tblPr_attrib(self.lastRow)

    @LastRow.setter
    def LastRow(self, value):
        self.set_tblPr_attrib(self.lastRow, value)

    @property
    def Parent(self):
        return self._shape

    @property
    def Rows(self):
        from .rows import Rows
        if self._rows is None:
            self._rows = Rows(self)
        return self._rows

    @property
    def VertBanding(self):
        return self.get_tblPr_attrib(self.bandCol)

    @VertBanding.setter
    def VertBanding(self, value):
        self.set_tblPr_attrib(self.bandCol, value)

    def get_tblPr_attrib(self, attr):
        return True if self.e_tblPr.get(attr, '0') == '1' else False

    def set_tblPr_attrib(self, attr, value):
        if value:
            self.e_tblPr.set(attr, '1')
        elif attr in self.e_tblPr.attrib:
            del self.e_tblPr.attrib[attr]
