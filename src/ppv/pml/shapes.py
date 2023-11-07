from ppv.pml.slide import Slide
from ppv.pml.slidelayout import SlideLayout
from ..core.collection import Collection
from PIL import Image
from ..img.image import Image as Img
from ..utils.length import Length


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

    def AddPicture(self, FileName, LinkToFile, SaveWithDocument, Left, Top,
                   Width=None, Height=None):

        return self._AddPicture(FileName, LinkToFile, SaveWithDocument,
                                Left, Top, Width, Height)

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

    def _AddPicture(self, FileName, LinkToFile, SaveWithDocument, Left, Top,
                    Width, Height):
        # observations:
        # - picture supported by Content, Picture, OnlinePicture placeholders
        # - preference is given to Picture placeholder to others
        # - When any of above placeholders are present these are used for pic
        # - Picture placeholder crops the picture if exceeds the ph size
        # - Content and Online picture placeholders fits the image inside
        # - crop is saved as t, b, l, r in ex. <a:srcRect t="17539" b="17539"/>
        # - a:srcRect must be next to a:blip
        # - <a:blip r:embed="rId2" r:link="rId3">
        # -     embed to save with doc
        # -     link to link to file
        # - x,y,cx,cy -> Left, Top, Width, Height
        # -     <a:xfrm>
        # -         <a:off x="422275" y="0"/>
        # -         <a:ext cx="3100387" cy="2952750"/>
        # -     </a:xfrm>
        # - ar = aspect ratio target
        # - ar_img = aspect ratio of image
        #
        # - sf = scale factor = w target / w source or h target / h source
        # - No crop: use w for scale factor if ar source >= ar target else h
        # - With crop: reverse of No crop
        #
        # - t = 100000 * crop / h
        # - l = 100000 * crop / w
        # -
        # - y = y0 + (h*sy*t/100000)
        # - x = x0 + (w*sx*l/100000)
        # - cy0 = cy + (h*sy*t/100000) + (h*sy*b/100000)
        # - cx0 = cx + (w*sx*l/100000) + (w*sx*r/100000)
        # - sx = cx0/w
        # - sy = cy0/h
        #
        # - x,y,cx,cy,t,b,l,r are integer not float
        #
        # - picture is added in place of placeholder then below is copied
        # -     <p:ph type="clipArt" sz="quarter" idx="15"/>
        #
        #
        # create p:pic element
        # index = -1
        # ph = relevant placeholder to replace
        # if ph
        #   index = index of ph
        #
        # insert pic at insert
        #
        # find a placeholder to replace picture with. First search for pic ph,
        # then search for content or online picture ph
        #
        # insert new pic element in place of placeholder or at end of spTree
        # if ph:
        #   move move the p:sp/p:nvSpPr/p:nvPr/p:ph to p:pic/p:nvPicPr/p:nvPr/
        # if ph is type pic:
        #     del children of p:pic/p:spPr
        #     calculate crop
        # if ph is type content/ online picture:
        #     DONT del children of p:pic/p:spPr
        #     calculate x y cx cy to fit the picture to ph
        # if no ph:
        #   calculate x y cx cy to fit the slide

        img = Image.open(FileName)
        w = Length.px2emu(img.width)
        h = Length.px2emu(img.height)
        ar_img = w / h

        part = self.Parent.part

        e = self.e.parse_file(self.e.data_path / "picture.xml").getroot()

        srcRect = e.find('.//'+e.qn('a:srcRect'))
        blip = e.find('.//'+e.qn('a:blip'))
        xfrm = e.find('.//'+e.qn('a:xfrm'))

        if SaveWithDocument:
            part_pic = part.package.add_part_picture(FileName)
            part_pic.typeobj.read(FileName)
            target_rel_uri_str = part.uri.get_rel(part_pic.uri.str)
            reltype = part_pic.typeobj.reltype
            rid = part.get_rels_part().typeobj.add_relation(reltype,
                                                            target_rel_uri_str)
            part.package.types.add_default(part_pic)
            blip.set(blip.qn('r:embed'), rid)
        else:
            del blip.attrib[blip.qn('r:embed')]

        if LinkToFile:
            target = 'file:///' + FileName
            reltype = Img.reltype
            rid2 = part.get_rels_part().typeobj.add_relation(reltype, target,
                                                             "External")
            blip.set(blip.qn('r:link'), rid2)
        else:
            del blip.attrib[blip.qn('r:link')]

        index = -1
        sp_ph = self.get_ph_to_add_picture()
        if sp_ph is not None:
            sp_ph, index, ph, is_pic_type = sp_ph
            e.find('.//'+e.qn('p:nvPr')).append(ph)

            try:
                cx = sp_ph.PictureFormat.cx
                cy = sp_ph.PictureFormat.cy
                x = sp_ph.PictureFormat.x
                y = sp_ph.PictureFormat.y
            except AttributeError:
                idx = ph.get('idx')
                sp_ph_ = self.get_shape_by_ph_idx(idx)
                cx = sp_ph_.PictureFormat.cx
                cy = sp_ph_.PictureFormat.cy
                x = sp_ph_.PictureFormat.x
                y = sp_ph_.PictureFormat.y

            ar = cx / cy
            if is_pic_type:
                sf = cy / h if ar <= ar_img else cx / w
                cx0 = sf * w
                cy0 = sf * h
                sx = cx0 / w
                sy = cy0 / h
                t = (cy0-cy) * 100000 / (h * sy)
                t, b = int(t/2), int(t/2)
                r = (cx0-cx) * 100000 / (w * sx)
                l, r = int(r/2), int(r/2)
                if t:
                    srcRect.set('t', str(t))
                if b:
                    srcRect.set('b', str(b))
                if l:
                    srcRect.set('l', str(l))
                if r:
                    srcRect.set('r', str(r))

                for i in e.findqn('p:spPr'):
                    i.getparent().remove(i)

            else:
                sf = cx / w if ar <= ar_img else cy / h
                cx_ = sf * w
                cy_ = sf * h
                dx = cx - cx_
                dy = cy - cy_
                x = int(x + dx / 2)
                y = int(y + dy / 2)
                cx = int(cx_)
                cy = int(cy_)

                off = xfrm.findqn('a:off')
                off.set('x', str(x))
                off.set('y', str(y))
                ext = xfrm.findqn('a:ext')
                ext.set('cx', str(cx))
                ext.set('cy', str(cy))
        else:
            x = Length.pt2emu(Left)
            y = Length.pt2emu(Top)
            sldSz = self.Parent.Parent.e.findqn('p:sldSz')
            cx_sld = int(sldSz.get('cx'))
            cy_sld = int(sldSz.get('cy'))

            if Width:
                cx = Length.pt2emu(Width)
            else:
                cx = cx_sld if w > cx_sld else w
            if Height:
                cy = Length.pt2emu(Height)
            else:
                cy = cy_sld if h > cy_sld else h

            ar = cx / cy
            sf = cx / w if ar <= ar_img else cy / h
            cx = int(sf * w)
            cy = int(sf * h)

            off = xfrm.findqn('a:off')
            off.set('x', str(x))
            off.set('y', str(y))
            ext = xfrm.findqn('a:ext')
            ext.set('cx', str(cx))
            ext.set('cy', str(cy))

        self.e.insert(index, e)
        if sp_ph:
            sp_ph.e.getparent().remove(sp_ph.e)

    def get_ph_to_add_picture(self):
        for i, sp in enumerate(self.Parent.Shapes.get_all()):
            e = sp.e
            if e.tag == e.qn('p:sp'):
                ph = e.find('.//'+e.qn('p:ph')+'[@type="pic"]')
                if ph is not None:
                    return sp, i+2, ph, True

        for i, sp in enumerate(self.Parent.Shapes.get_all()):
            e = sp.e
            if e.tag == e.qn('p:sp'):
                ph = e.find('.//'+e.qn('p:ph'))
                if ph is not None:
                    a = ph.attrib
                    if ('type' not in a) or (a.get('type') == 'clipArt'):
                        return sp, i+2, ph, False

    def get_shape_by_ph_idx(self, idx):
        if isinstance(self.Parent, Slide):
            sl = self.Parent.CustomLayout
        elif isinstance(self.Parent, SlideLayout):
            sl = self.Parent
        else:
            return

        for sp in sl.Shapes.get_all():
            ph = sp.e.find('.//'+sp.e.qn('p:ph')+'[@idx="'+idx+'"]')
            if ph is not None:
                return sp
