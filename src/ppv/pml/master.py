
from ..core.xmltypeobj import XmlTypeobj


class Master(XmlTypeobj):
    type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"

    @property
    def CustomLayouts(self):
        from ppv.pml.customlayouts import CustomLayouts
        if not hasattr(self, "_customlayouts"):
            self._customlayouts = CustomLayouts(self)
        return self._customlayouts
