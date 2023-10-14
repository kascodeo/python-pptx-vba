from .font import Font


class Font2(Font):

    @property
    def Fill(self):
        from ppv.dml.fillformat import FillFormat
        if not hasattr(self, '_fill'):
            self._fill = FillFormat(self)
        return self._fill
