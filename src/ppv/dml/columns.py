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

    def Add(self):
        pass
