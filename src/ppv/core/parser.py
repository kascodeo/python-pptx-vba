from lxml import etree
from opc.parser import ElementBase as _ElementBase


class ElementBase(_ElementBase):
    def findqn(self, path, namespaces=None):
        path = self.qn(path, nsmap=namespaces)
        return super().find(path)

    def findallqn(self, path, namespaces=None):
        path = self.qn(path, nsmap=namespaces)
        return super().findall(path)

    def getqn(self, key, default=None):
        key = self.qn(key)
        return super().get(key, default)


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
