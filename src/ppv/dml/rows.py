from ..core.collection import Collection


class Rows(Collection):
    def __init__(self, table):
        super().__init__(table)
        self._table = table
        self._coll = {}

    @property
    def e(self):
        return self.Parent.e

    @property
    def e_lst_tr(self):
        return self.e.findallqn('a:tr')

    def get_all(self):
        from .row import Row
        lst = []
        for e in self.e.findallqn('a:tr'):
            if e not in self._coll:
                self._coll[e] = Row(self.Parent, e)
            lst.append(self._coll[e])
        return lst

    def Add(self, BeforeRow=-1):
        if BeforeRow not in [-1] + list(range(1, self.Count+1)):
            raise ValueError("BeforeRow must be from 1 to Count(inclusive)")
        index_to_copy = 0 if BeforeRow == 1 else \
            self.Count - 1 if BeforeRow == -1 else \
            BeforeRow - 2

        e_row_to_copy = self.get_all()[index_to_copy].e
        e_row_new = e_row_to_copy.deepcopy()

        if BeforeRow == -1:
            e_row_to_copy.addnext(e_row_new)
        else:
            self.get_all()[BeforeRow-1].e.addprevious(e_row_new)

        for i, r in enumerate(self.e_lst_tr):
            if e_row_new is r:
                row = self.Item(i+1)
                row.change_rowId()
                for cell in row.Cells.get_all():
                    cell.clear_text()
                return row
