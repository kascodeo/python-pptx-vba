class Row():
    def __init__(self, table, e):
        self._table = table
        self._e = e  # 'a:tr'
        self._cells = {}
        self._cellrange = None

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
        from .cellrange import CellRange
        if self._cellrange is None:
            self._cellrange = CellRange(self)
        return self._cellrange

    @property
    def Height(self):
        pass

    @property
    def index(self):
        for i, e in enumerate(self.Parent.Rows.e_lst_tr):
            if e is self.e:
                return i

    @property
    def rowId(self):
        e = self.e
        return e.find('./'+e.qn('a:extLst')+'/'+e.qn('a:ext')+'/*[@val]').\
            get('val')

    def change_rowId(self):
        from ppv.utils.randomid import RandomId
        id = str(RandomId.get())
        e = self.e
        e.find('./'+e.qn('a:extLst')+'/'+e.qn('a:ext')+'/*[@val]').\
            set('val', id)
