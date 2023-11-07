from ..core.imagetypeobj import ImageTypeobj


class Image(ImageTypeobj):
    types = [
        "image/bmp",
        "image/gif",
        "image/ico",
        "image/jpeg",
        "image/jpg",
        "image/png",
        "image/tiff",
        "image/webp",
        "image/wmf",
        "image/emf",

    ]
    reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
