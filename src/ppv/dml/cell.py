class Cell():
    def __init__(self, table, e):
        self._table = table
        self._e = e
        self._shape = None

    @property
    def e(self):  # 'a:tc'
        return self._e

    def Merge(self):
        pass

    def Split(self):
        pass

    @property
    def Borders(self):
        pass

    @property
    def Parent(self):
        return self._table

    @property
    def Shape(self):
        from ..pml.shape import Shape
        if self._shape is None:
            self._shape = Shape(self, self.e)
        return self._shape

    def clear_text(self):
        e_cell = self.e
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
