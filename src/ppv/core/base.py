from abc import ABC, abstractmethod


class AppMixin():
    @property
    def Application(self):
        from ppv.pml.application import Application
        return Application()


class TypeobjBase(ABC):
    """Abstract class. Extending classes must implement read, write methods"""

    def __init__(self, part):
        self._part = part

    @property
    def part(self):
        return self._part

    @abstractmethod
    def read(self):
        pass

    @abstractmethod
    def write(self, f):
        pass
