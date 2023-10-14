from ..core.collection import Collection


class CustomLayouts(Collection):

    def get_all(self):
        from ppv.pml.slidelayout import SlideLayout
        parts = self.Parent.part.get_related_parts_by_reltype(
            SlideLayout.reltype)
        parts.sort(key=lambda part: int(part.uri.str.replace(
            '/ppt/slideLayouts/slideLayout', '').replace('.xml', '')))

        return [part.typeobj for part in parts]
