from enum import Enum


class MsoTextOrientation(Enum):
    msoTextOrientationDownward = 3  # Downward
    msoTextOrientationHorizontal = 1  # Horizontal
    # Horizontal and rotated as required for Asian language support
    msoTextOrientationHorizontalRotatedFarEast = 6
    msoTextOrientationMixed = -2  # Not supported
    msoTextOrientationUpward = 2  # Upward
    msoTextOrientationVertical = 5  # Vertical
    # Vertical as required for Asian language support
    msoTextOrientationVerticalFarEast = 4
