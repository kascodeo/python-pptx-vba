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
        lst = [rpr.get('b', '0') for rpr in self.Parent.get_rpr()]
        return MsoTriState.msoTrue if all(map(lambda i: i == '1', lst)) else \
            MsoTriState.msoFalse if all(map(lambda i: i == '0', lst)) else \
            MsoTriState.msoTriStateMixed

    @Bold.setter
    def Bold(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        self.Parent.isolate()
        if val in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            [rpr.set('b', '1') for rpr in self.Parent.get_rpr()]
        else:
            for rpr in self.Parent.get_rpr():
                if 'b' in rpr.attrib:
                    del rpr.attrib['b']

    @property
    def Italic(self):
        lst = [rpr.get('i', '0') for rpr in self.Parent.get_rpr()]
        return MsoTriState.msoTrue if all(map(lambda i: i == '1', lst)) else \
            MsoTriState.msoFalse if all(map(lambda i: i == '0', lst)) else \
            MsoTriState.msoTriStateMixed

    @Italic.setter
    def Italic(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        self.Parent.isolate()
        if val in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            [rpr.set('i', '1') for rpr in self.Parent.get_rpr()]
        else:
            for rpr in self.Parent.get_rpr():
                if 'i' in rpr.attrib:
                    del rpr.attrib['i']

    @property
    def Name(self):
        lst = []
        for rpr in self.Parent.get_rpr():
            latin = rpr.findqn('a:latin')
            name = self.default_name
            if latin is not None:
                name = latin.get('typeface', "")
            lst.append(name)
        return lst[0] if len(set(lst)) == 1 else ""

    @Name.setter
    def Name(self, val):
        self.Parent.isolate()
        for rpr in self.Parent.get_rpr():
            latin = rpr.findqn('a:latin')
            if latin is None:
                latin = rpr.makeelement(rpr.qn('a:latin'))
                rpr.append(latin)
            latin.set('typeface', val)

    @property
    def Size(self):
        lst = [rpr.get('sz', self.default_sz) for rpr in self.Parent.get_rpr()]
        return float(lst[0]) / 100 if len(set(lst)) == 1 else ""

    @Size.setter
    def Size(self, val):
        self.Parent.isolate()
        val = float(val)*100
        if float(val) != float(self.default_sz):
            val = '{:g}'.format(float(val))
            [rpr.set('sz', val) for rpr in self.Parent.get_rpr()]
        else:
            for rpr in self.Parent.get_rpr():
                if 'sz' in rpr.attrib:
                    del rpr.attrib['sz']

    @property
    def Underline(self):
        lst = [rpr.get('u', '') for rpr in self.Parent.get_rpr()]
        return MsoTriState.msoTrue if all(map(lambda i: i == 'sng', lst)) else\
            MsoTriState.msoFalse if all(map(lambda i: i == '', lst)) else \
            MsoTriState.msoTriStateMixed

    @Underline.setter
    def Underline(self, val):
        if val not in [MsoTriState.msoTrue, MsoTriState.msoTrue.value,
                       MsoTriState.msoFalse, MsoTriState.msoFalse.value,
                       ]:
            raise ValueError("value must be msoTrue|msoFalse|their values")
        self.Parent.isolate()
        if val in [MsoTriState.msoTrue, MsoTriState.msoTrue.value]:
            [rpr.set('u', 'sng') for rpr in self.Parent.get_rpr()]
        else:
            for rpr in self.Parent.get_rpr():
                if 'u' in rpr.attrib:
                    del rpr.attrib['u']
