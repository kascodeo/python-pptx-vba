from ..utils.length import Length


class PictureFormat():
    CC = 100000  # crop constant

    def __init__(self, parent):
        self._parent = parent

    @property
    def Parent(self):
        return self._parent

    @property
    def Crop(self):
        pass

    @property
    def CropBottom(self):
        pass

    @CropBottom.setter
    def CropBottom(self, value):
        # add b attr in srcRect
        # cy attr of a:ext
        cb = Length.pt2emu(value)

        sy = self.sy
        h = self.h
        cy0 = self.cy0
        ct = self.ct

        b = int(cb / h * self.CC)
        cy = int(cy0 - (ct + cb) * sy)

        srcRect = self.get_or_add_srcRect()  # creates element if absent

        srcRect.set('b', str(b))
        self.e_ext.set('cy', str(cy))

    @property
    def CropLeft(self):
        pass

    @CropLeft.setter
    def CropLeft(self, value):
        # add t attr in srcRect
        # change y attr of a:off and cy attr of a:ext

        cl = Length.pt2emu(value)

        sx = self.sx
        w = self.w
        x0 = self.x0
        cx0 = self.cx0
        cr = self.cr

        l_ = int(cl / w * self.CC)
        x = int(x0 + cl*sx)
        cx = int(cx0 - (cl + cr) * sx)

        srcRect = self.get_or_add_srcRect()  # creates element if absent

        srcRect.set('l', str(l_))
        self.e_off.set('x', str(x))
        self.e_ext.set('cx', str(cx))

    @property
    def CropRight(self):
        pass

    @CropRight.setter
    def CropRight(self, value):
        # add r attr in srcRect
        # cx attr of a:ext
        cr = Length.pt2emu(value)

        sx = self.sx
        w = self.w
        cx0 = self.cx0
        cl = self.cl

        r = int(cr / w * self.CC)
        cx = int(cx0 - (cl + cr) * sx)

        srcRect = self.get_or_add_srcRect()  # creates element if absent

        srcRect.set('r', str(r))
        self.e_ext.set('cx', str(cx))

    @property
    def CropTop(self):
        # return self.
        pass

    @CropTop.setter
    def CropTop(self, value):
        # add t attr in srcRect
        # change y attr of a:off and cy attr of a:ext

        ct = Length.pt2emu(value)

        sy = self.sy
        h = self.h
        y0 = self.y0
        cy0 = self.cy0
        cb = self.cb

        t = int(ct / h * self.CC)
        y = int(y0 + ct*sy)
        cy = int(cy0 - (ct + cb) * sy)

        srcRect = self.get_or_add_srcRect()  # creates element if absent

        srcRect.set('t', str(t))
        self.e_off.set('y', str(y))
        self.e_ext.set('cy', str(cy))

    @property
    def cb(self):
        return int(self.b*self.h / self.CC)

    @property
    def ct(self):
        return int(self.t*self.h / self.CC)

    @property
    def cl(self):
        return int(self.l_*self.w / self.CC)

    @property
    def cr(self):
        return int(self.r*self.w / self.CC)

    @property
    def imgtypeobj(self):
        part = self.Parent.Parent.part.get_related_part(self.rid)
        return part.typeobj

    @property
    def img(self):
        return self.imgtypeobj.img

    @property
    def x(self):
        return int(self.e_off.get('x'))

    @property
    def y(self):
        return int(self.e_off.get('y'))

    @property
    def cx(self):
        return int(self.e_ext.get('cx'))

    @property
    def cy(self):
        return int(self.e_ext.get('cy'))

    @property
    def t(self):
        return 0 if self.e_srcRect is None else int(self.e_srcRect.get('t', 0))

    @property
    def b(self):
        return 0 if self.e_srcRect is None else int(self.e_srcRect.get('b', 0))

    @property
    def l_(self):
        return 0 if self.e_srcRect is None else int(self.e_srcRect.get('l', 0))

    @property
    def r(self):
        return 0 if self.e_srcRect is None else int(self.e_srcRect.get('r', 0))

    @property
    def w(self):
        return Length.px2emu(self.img.width)

    @property
    def h(self):
        return Length.px2emu(self.img.height)

    @property
    def sx(self):
        return self.cx / (self.w*(1-(self.l_/self.CC) - (self.r/self.CC)))

    @property
    def sy(self):
        return self.cy / (self.h*(1-(self.t/self.CC) - (self.b/self.CC)))

    @property
    def y0(self):
        return self.y - (self.t / self.CC) * self.h * self.sy

    @property
    def x0(self):
        return self.x - (self.l_ / self.CC) * self.w * self.sx

    @property
    def cx0(self):
        return self.cx + (self.l_ + self.r) / self.CC * self.w * self.sx

    @property
    def cy0(self):
        return self.cy + (self.t + self.b) / self.CC * self.h * self.sy

    @property
    def e_xfrm(self):
        e = self.Parent.e
        return e.find('.//'+e.qn('a:xfrm'))

    @property
    def e_off(self):
        return self.e_xfrm.findqn('a:off')

    @property
    def e_ext(self):
        return self.e_xfrm.findqn('a:ext')

    def get_or_add_srcRect(self):
        e = self.Parent.e
        if e.ln == 'pic':
            if self.e_srcRect is None:
                e = self.e_blipFill.makeelement(e.qn('a:srcRect'))
                self.e_blip.addnext(e)
                # self.e_blipFill.append(e)
            return self.e_srcRect

    @property
    def e_blipFill(self):
        return self.Parent.e.findqn('p:blipFill')

    @property
    def rid(self):
        return self.e_blip.get(self.e_blip.qn('r:embed'))

    @property
    def e_blip(self):
        return self.e_blipFill.findqn('a:blip')

    @property
    def e_srcRect(self):
        return self.e_blipFill.findqn('a:srcRect')
