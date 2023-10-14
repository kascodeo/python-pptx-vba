from .textrange import TextRange


class TextRange2(TextRange):

    @property
    def Font(self):
        from ppv.dml.font2 import Font2
        return Font2(self)
