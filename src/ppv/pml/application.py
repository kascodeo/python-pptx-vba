class Application():
    """Object of this class represents the Powerpoint Application in this
    library"""
    __instance = None

    def __new__(cls):
        # Singleton implementation
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self):
        # initiates the properties
        from .presentations import Presentations

        if not hasattr(self.__instance, '_presentations'):
            self._presentations = Presentations(self)

    @property
    def Presentations(self):
        """Returns the Presentations collection object"""
        return self._presentations
