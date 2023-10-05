from .base import AppMixin


class Collection(AppMixin):
    def __init__(self, parent):
        self._parent = parent

    @property
    def Parent(self):
        return self._parent
