from ..core.collection import Collection


class CellRange(Collection):
    def __init__(self, parent):  # parent is Row|Column
        super().__init__(parent)
        self._coll = {}

    @property
    def e(self):
        return self.Parent.e  # 'a:tr' for Row, 'a:gridCol' for Column

    def get_all(self):
        from .column import Column
        from .row import Row
        from .cell import Cell
        if isinstance(self.Parent, Column):
            col_index = self.Parent.index
            table = self.Parent.Parent.Parent.Table
            lst = []
            for row in table.Rows.get_all():
                lst.append(row.Cells.Item(col_index+1))
            return lst

        elif isinstance(self.Parent, Row):
            lst = []
            for e in self.e.findallqn('a:tc'):
                if e not in self._coll:
                    self._coll[e] = Cell(self.Parent.Parent, e)
                lst.append(self._coll[e])
            return lst

    @property
    def Borders(self):
        pass
