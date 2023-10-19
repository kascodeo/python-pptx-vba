from ppv.enum.MsoTriState import MsoTriState


class Font():
    default_name = "Calibri"
    default_sz = "1800"

    def __init__(self, textrange):
        self._textrange = textrange

    @property
    def Parent(self):
        return self._textrange

    @property
    def Bold(self):
        lst = [rpr.get('b', '0') for rpr in self.Parent.get_rpr_in_range()]
        return MsoTriState.msoTrue if all(map(lambda i: i == '1', lst)) else \
            MsoTriState.msoFalse if all(map(lambda i: i == '0', lst)) else \
            MsoTriState.msoTriStateMixed

    @Bold.setter
    def Bold(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        attr = 'b'
        value = '1'
        self.Parent.isolate()
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            value = None
        for rpr in self.Parent.get_rpr_in_range():
            self.Parent.update_rpr(rpr, attr, value)

    @property
    def Italic(self):
        lst = [rpr.get('i', '0') for rpr in self.Parent.get_rpr_in_range()]
        return MsoTriState.msoTrue if all(map(lambda i: i == '1', lst)) else \
            MsoTriState.msoFalse if all(map(lambda i: i == '0', lst)) else \
            MsoTriState.msoTriStateMixed

    @Italic.setter
    def Italic(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        attr = 'i'
        value = '1'
        self.Parent.isolate()
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            value = None
        for rpr in self.Parent.get_rpr_in_range():
            self.Parent.update_rpr(rpr, attr, value)

    @property
    def Name(self):
        lst = []
        for rpr in self.Parent.get_rpr_in_range():
            latin = rpr.findqn('a:latin')
            name = self.default_name
            if latin is not None:
                name = latin.get('typeface', "")
            lst.append(name)
        return lst[0] if len(set(lst)) == 1 else ""

    @Name.setter
    def Name(self, val):
        self.Parent.isolate()
        attr = 'font'
        value = val
        if val == self.default_name:
            value = None
        for rpr in self.Parent.get_rpr_in_range():
            self.Parent.update_rpr(rpr, attr, value)

    @property
    def Size(self):
        lst = [rpr.get('sz', self.default_sz)
               for rpr in self.Parent.get_rpr_in_range()]
        return self.get_size(lst[0]) if len(set(lst)) == 1 else ""

    @Size.setter
    def Size(self, val):
        self.Parent.isolate()
        attr = 'sz'
        value = '{:g}'.format(float(val)*100)
        if float(value) == float(self.default_sz):
            value = None
        for rpr in self.Parent.get_rpr_in_range():
            self.Parent.update_rpr(rpr, attr, value)

    @property
    def Underline(self):
        lst = [rpr.get('u', '') for rpr in self.Parent.get_rpr_in_range()]
        return MsoTriState.msoTrue if all(map(lambda i: i == 'sng', lst)) else\
            MsoTriState.msoFalse if all(map(lambda i: i == '', lst)) else \
            MsoTriState.msoTriStateMixed

    @Underline.setter
    def Underline(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        attr = 'u'
        value = 'sng'
        self.Parent.isolate()
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            value = None
        for rpr in self.Parent.get_rpr_in_range():
            self.Parent.update_rpr(rpr, attr, value)

    def get_size(self, sz):
        return '{:g}'.format(float(sz) / 100)
