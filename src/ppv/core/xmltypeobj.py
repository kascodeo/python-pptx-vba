from lxml import etree
from .base import TypeobjBase
from .base import AppMixin


class XmlTypeobj(AppMixin, TypeobjBase):
    """Must implement read write methods as per python-opc-lite library

    """
    @property
    def e(self):
        return self._e

    @e.setter
    def e(self, value):
        self._e = value

    @property
    def Parent(self):
        return self.part.package.document

    @property
    def parser(self):
        from .parser import Parser
        return Parser()

    def read(self, f):
        """parse the given file object as xml and assigns the root element to
        self._e
        """
        self._e = self.parser.parse(f).getroot()

    def write(self, f):
        """writes the xml content in self._e to the given file object. utf-8
        encoding, xml_declaration and standalone properties are applied to xml
        """
        f.write(etree.tostring(self.e, encoding='UTF-8',
                               xml_declaration=True, standalone=True))

    def get_related_typeobj(self, rid):
        return self.part.get_related_part(rid).typeobj
