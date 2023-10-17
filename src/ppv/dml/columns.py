from ..core.collection import Collection


class Columns(Collection):
    def __init__(self, table):
        super().__init__(table)
        self._table = table
        self._coll = {}

    @property
    def e(self):
        return self.Parent.e

    @property
    def e_tblGrid(self):
        return self.Parent.e.findqn('a:tblGrid')

    @property
    def e_lst_gridCol(self):
        return self.e_tblGrid.findallqn('a:gridCol')

    def get_all(self):
        from .column import Column
        lst = []
        for e in self.e_tblGrid.findallqn('a:gridCol'):
            if e not in self._coll:
                self._coll[e] = Column(self.Parent, e)
            lst.append(self._coll[e])
        return lst

    def Add(self, BeforeColumn=-1):
        if BeforeColumn not in [-1] + list(range(1, self.Count+1)):
            raise ValueError("BeforeColumn must be from 1 to Count(inclusive)")

        index_to_copy = 0 if BeforeColumn == 1 else \
            self.Count - 1 if BeforeColumn == -1 else \
            BeforeColumn - 2

        e_col_to_copy = self.get_all()[index_to_copy].e
        e_col_new = e_col_to_copy.deepcopy()

        index_insert = self.Count if BeforeColumn == -1 else BeforeColumn - 1
        self.e_tblGrid.insert(index_insert, e_col_new)

        for row in self.Parent.Rows.get_all():
            cell_to_copy = row.Cells.Item(index_to_copy+1)
            e_cell_new = cell_to_copy.e.deepcopy()
            e_cell = e_cell_new
            for p in e_cell.findqn('a:txBody').findallqn('a:p')[1:]:
                p.getparent().remove(p)
            p = e_cell.findqn('a:txBody').findqn('a:p')
            if p is not None:
                for r in p.findallqn('a:r')[1:]:
                    r.getparent().remove(r)
                r = p.findqn('a:r')
                if r is not None:
                    t = r.findqn('a:t')
                    if t is not None:
                        t.text = ''

            if BeforeColumn == -1:
                row.Cells.get_all()[-1].e.addnext(e_cell_new)
            else:
                row.Cells.get_all()[BeforeColumn-1].e.addprevious(e_cell_new)

        for col in self.get_all():
            if col.e is e_col_new:
                col.change_colId()
                for cell in col.Cells.get_all():
                    cell.clear_text()
                return col
