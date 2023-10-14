from .textframe import TextFrame


class TextFrame2(TextFrame):
    @property
    def TextRange(self):
        from ppv.dml.textrange2 import TextRange2
        return TextRange2(self)
