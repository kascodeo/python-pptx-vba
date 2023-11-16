from lxml import etree
from copy import deepcopy
from pathlib import Path
from opc.parser import ElementBase as _ElementBase


class ElementBase(_ElementBase):

    @property
    def ln(self):
        return etree.QName(self).localname

    def findqn(self, path, namespaces=None):
        path = self.qn(path, nsmap=namespaces)
        return super().find(path)

    def findallqn(self, path, namespaces=None):
        path = self.qn(path, nsmap=namespaces)
        return super().findall(path)

    def getqn(self, key, default=None):
        key = self.qn(key)
        return super().get(key, default)

    def setqn(self, key, value, nsmap=None):
        super().set(self.qn(key, nsmap), value)

    def deepcopy(self):
        return deepcopy(self)

    @property
    def makeelement(self):
        return self.parser.makeelement

    @property
    def data_path(self):
        return Path(__file__).parent.parent / "data"

    @property
    def parser(self):
        return Parser()

    def parse_file(self, path):
        with open(path) as f:
            return self.parser.parse(f)

    @property
    def length(self):
        from ppv.utils.length import Length
        return Length


class Parser(etree.XMLParser):
    __instance = None

    def __new__(cls):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self):
        self.set_element_class_lookup(
            etree.ElementDefaultClassLookup(ElementBase))

    def parse(self, fp):
        return etree.parse(fp, self)
