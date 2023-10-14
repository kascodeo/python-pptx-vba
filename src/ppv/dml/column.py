class Column():
    def __init__(self, table, e):
        self._table = table
        self._e = e  # 'a:gridCol'

    @property
    def e(self):
        return self._e

    @property
    def Parent(self):
        return self._table

    def Delete(self):
        pass

    @property
    def Cells(self):
        pass

    @property
    def Width(self):
        pass

    @property
    def index(self):
        for i, e in enumerate(self.Parent.Columns.e_lst_gridCol):
            if e is self.e:
                return i
