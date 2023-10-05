from ..core.collection import Collection


class Slides(Collection):

    @property
    def e(self):
        return self.Parent.e.findqn('p:sldIdLst')

    def get_all(self):
        return self.e.findallqn('p:sldId')

    @property
    def Count(self):
        return len(self.get_all())

    def AddSlide(self, Index, pCustomLayout):
        pass

    def FindBySlideID(self, SlideID):
        if not isinstance(SlideID, int):
            raise TypeError("SlideID has to be int")

        for sldId in self.get_all():
            if int(sldId.get('id')) == SlideID:
                rid = sldId.getqn('r:id')
                return self.Parent.get_related_typeobj(rid)

    def Item(self, Index):
        if not isinstance(Index, int):
            raise TypeError("Index has to be int")

        if (Index < 1) or (Index > self.Count):
            raise IndexError("Out of Range Index")

        rid = self.get_all()[Index - 1].getqn('r:id')
        return self.Parent.get_related_typeobj(rid)
