from abc import abstractmethod
from .base import AppMixin


class Collection(AppMixin):
    def __init__(self, parent):
        self._parent = parent

    @property
    def Parent(self):
        return self._parent

    @property
    def Count(self):
        return len(self.get_all())

    @abstractmethod
    def get_all(self):
        pass

    def Item(self, Index):
        if not isinstance(Index, int):
            raise TypeError("Index has to be int")

        if (Index < 1) or (Index > self.Count):
            raise IndexError("Out of Range Index")

        return self.get_all()[Index - 1]
