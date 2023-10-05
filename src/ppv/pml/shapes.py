from ..core.collection import Collection


class Shapes(Collection):
    @property
    def e(self):
        return self.Parent.e.find('.//' + self.Parent.e.qn('p:spTree'))

    def get_all(self):
        lst = []
        for e in self.e:
            if e.tag not in [self.e.qn('p:nvGrpSpPr'), self.e.qn('p:grpSpPr')]:
                lst.append(e)
        return lst

    @property
    def Count(self):
        return len(self.get_all())

    @property
    def Placeholders(self):
        from ppv.pml.placeholders import Placeholders
        if not hasattr(self, "_placeholders"):
            self._placeholders = Placeholders(self.Parent)
        return self._placeholders
