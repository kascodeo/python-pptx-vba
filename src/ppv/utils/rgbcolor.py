class MsoRGBType():
    def __init__(self, r, g, b):
        for i in [r, g, b]:
            if not isinstance(i, int):
                raise TypeError("color values must be integer type")
            if not (0 <= i < 256):
                raise ValueError("color must be from 0 to 255")

        self.r = r
        self.g = g
        self.b = b

    @classmethod
    def fromhex(cls, hex):
        hex = hex.lstrip('#')
        return cls(int(hex[0:2]), int(hex[2:4]), int(hex[4:6]))

    @property
    def hexstr(self):
        return '{:02x}{:02x}{:02x}'.format(self.r, self.g, self.b)


def RGB(r, g, b):
    return MsoRGBType(r, g, b)
