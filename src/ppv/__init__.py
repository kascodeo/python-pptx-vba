__version__ = '0.0.1'

from .pml.application import Application as _Application

Application = _Application()
Presentations = Application.Presentations
