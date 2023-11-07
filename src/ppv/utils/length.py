class Length():
    PT = 12700      # emu
    CP = 127        # emu
    MM = 36000      # emu
    CM = 360000     # emu
    INCH = 914400   # emu
    PX = 9525       # emu

    @classmethod
    def pt2emu(cls, pt):
        return pt * cls.PT

    @classmethod
    def px2emu(cls, pt):
        return pt * cls.PX
