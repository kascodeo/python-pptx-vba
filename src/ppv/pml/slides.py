from ..core.collection import Collection


class Slides(Collection):

    @property
    def package(self):
        return self.Parent.part.package

    @property
    def e(self):
        return self.Parent.e.findqn('p:sldIdLst')

    def get_all(self):
        return [self.Parent.get_related_typeobj(sldId.getqn('r:id'))
                for sldId in self.e]

    def AddSlide(self, Index, pCustomLayout):
        from ppv.pml.slidelayout import SlideLayout
        from ppv.pml.slide import Slide

        if not isinstance(pCustomLayout, SlideLayout):
            raise ValueError("pCustomLayout must be of type SlideLayout")

        # create p:sld element
        sld = self.e.makeelement(self.e.qn('p:sld'), nsmap=self.e.nsmap)

        # copy to p:sld, each child of p:sldLayout element of pCustomLayout
        for c in pCustomLayout.e:
            sld.append(c.deepcopy())

        # remove name of cSld
        if 'name' in sld[0].attrib:
            sld[0].attrib.pop('name')

        # delete p:sp that has:
        #   p:nvSpPr/p:nvPr/p:ph[@type="dt"]
        #   p:nvSpPr/p:nvPr/p:ph[@type="ftr"]
        #   p:nvSpPr/p:nvPr/p:ph[@type="sldNum"]
        for sp in sld.findall('.//'+sld.qn('p:sp')):
            xpaths_todel = ['.//'+sp.qn('p:ph')+'[@type="dt"]',
                            './/'+sp.qn('p:ph')+'[@type="ftr"]',
                            './/'+sp.qn('p:ph')+'[@type="sldNum"]',
                            ]
            for xpath in xpaths_todel:
                if sp.find(xpath) is not None:
                    sp.getparent().remove(sp)

        # remove the children of p:spPr of all p:sp
        for sp in sld.findall('.//'+sld.qn('p:sp')):
            spPr = sp.findqn('p:spPr')
            if spPr is not None:
                for c in spPr:
                    c.getparent().remove(c)

            # remove the a:r of p:sp/p:txBody/a:p
            txBody = sp.findqn('p:txBody')
            if txBody is not None:
                for p in txBody.findallqn('a:p')[:-1]:
                    p.getparent().remove(p)
                p = txBody.findqn('a:p')
                pPr = p.findqn('a:pPr')
                if pPr is not None:
                    pPr.getparent().remove(pPr)
                for r in p.findallqn('a:r'):
                    r.getparent().remove(r)

        # Add part in package, add typeobj to part
        part = self.package.add_part(self.get_next_uri_str(), Slide.type)

        # add type in types of package
        self.package.types.add_override(part)

        # add p:sld to typeobj as _e
        part.typeobj.e = sld

        # add relspart in package; add typeobj to relspart
        relspart = part.add_relspart()

        # add relationship of customlayout to typeobj of relspart
        target_rel_uri_str = part.uri.get_rel(pCustomLayout.part.uri.str)
        relspart.typeobj.add_relation(
            pCustomLayout.reltype, target_rel_uri_str)

        # add part and relspart to types of package
        self.package.types.add_override(part)
        self.package.types.add_default(relspart)

        # add to presentation's sldIdLst
        target = self.Parent.part.uri.get_rel(part.uri.str)
        rid = self.Parent.part.get_rels_part().typeobj.add_relation(
            Slide.reltype, target)

        attrib = {'id': self.get_next_id(), self.e.qn('r:id'): rid}
        sldId = self.e.makeelement(self.e.qn('p:sldId'), attrib=attrib)
        self.e.insert(Index-1, sldId)

        return part.typeobj

    def FindBySlideID(self, SlideID):
        if not isinstance(SlideID, int):
            raise TypeError("SlideID has to be int")

        for sldId in self.e:
            if int(sldId.get('id')) == SlideID:
                rid = sldId.getqn('r:id')
                return self.Parent.get_related_typeobj(rid)

    def get_next_uri_str(self):
        existing = [slide.part.uri.str for slide in self.get_all()]
        i = 1
        while True:
            uri_str = "/ppt/slides/slide"+str(i)+".xml"
            if uri_str not in existing:
                return uri_str
            i += 1

    def get_next_id(self):
        used = 256
        for sldId in self.e:
            id = int(sldId.get('id'))
            if id > used:
                used = id
        return str(used + 1)
