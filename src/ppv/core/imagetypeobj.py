from .base import TypeobjBase
from PIL import Image


class ImageTypeobj(TypeobjBase):
    @property
    def img(self):
        return self._img

    def read(self, f):
        self._img = Image.open(f)
        self._img.load()

    def write(self, f):
        ext = self.part.uri.ext
        self._img.save(f, ext)
