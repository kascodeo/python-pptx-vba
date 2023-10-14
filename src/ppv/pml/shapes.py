from ..core.collection import Collection


class Shapes(Collection):
    def __init__(self, parent):
        super().__init__(parent)
        self._coll = {}

    @property
    def e(self):
        return self.Parent.e.find('.//' + self.Parent.e.qn('p:spTree'))

    def get_all(self):
        from ppv.pml.shape import Shape
        lst = []
        for e in self.e:
            if e.tag not in [self.e.qn('p:nvGrpSpPr'), self.e.qn('p:grpSpPr')]:
                if e not in self._coll:
                    self._coll[e] = Shape(self, e)
                lst.append(self._coll[e])
        return lst

    @property
    def Placeholders(self):
        from ppv.pml.placeholders import Placeholders
        if not hasattr(self, "_placeholders"):
            self._placeholders = Placeholders(self.Parent)
        return self._placeholders

    def AddTextbox(self, Orientation, Left, Top, Width, Height):
        from ppv.enum.MsoTextOrientation import MsoTextOrientation as mto

        e = self.e.parse_file(self.e.data_path / "textbox.xml").getroot()
        bodyPr = e.find('.//'+self.e.qn('a:bodyPr'))

        orient_dict = {
            mto.msoTextOrientationDownward: 'vert',
            mto.msoTextOrientationDownward.value: 'vert',

            mto.msoTextOrientationHorizontal: 'horz',
            mto.msoTextOrientationHorizontal.value: 'horz',

            mto.msoTextOrientationHorizontalRotatedFarEast: 'wordArtVert',
            mto.msoTextOrientationHorizontalRotatedFarEast.value:
                'wordArtVert',

            mto.msoTextOrientationUpward: 'vert270',
            mto.msoTextOrientationUpward.value: 'vert270',

            mto.msoTextOrientationVertical: 'mongolianVert',
            mto.msoTextOrientationVertical.value: 'mongolianVert',

            mto.msoTextOrientationVerticalFarEast: 'eaVert',
            mto.msoTextOrientationVerticalFarEast.value: 'eaVert',
        }
        if Orientation not in orient_dict:
            raise ValueError("Orientation is not supported yet.")

        bodyPr.set('vert', orient_dict[Orientation])
        off = e.find('.//'+e.qn('a:off'))
        ext = e.find('.//'+e.qn('a:xfrm')+'/'+e.qn('a:ext'))
        off.set('x', str(e.length.pt2emu(Left)))
        off.set('y', str(e.length.pt2emu(Top)))
        ext.set('cx', str(e.length.pt2emu(Width)))
        ext.set('cy', str(e.length.pt2emu(Height)))

        e.find('.//'+e.qn('p:cNvPr')).set('id', self.get_next_id())
        self.e.append(e)
        return self.Item(self.Count)

    def get_next_id(self):
        return str(max([int(e.get('id')) for e in
                        self.e.findall('.//'+self.e.qn('p:cNvPr')+'[@id]')
                        ]) + 1)

    def AddChart2(self):
        pass

    def AddConnector(self):
        pass

    def AddPicture(self):
        pass

    def AddPicture2(self):
        pass

    def AddTable(self):
        pass

    def AddOLEObject(self):
        pass

    def AddMediaObject(self):
        pass

    def AddMediaObject2(self):
        pass

    def AddShape(self):
        pass

    def BuildFreeform(self):
        pass

    def HasTitle(self):
        pass

    def Title(self):
        pass
