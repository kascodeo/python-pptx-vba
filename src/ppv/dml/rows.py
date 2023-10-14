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

    def Add(self):
        pass
