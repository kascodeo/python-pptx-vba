from pathlib import Path
from .presentation import Presentation


class Presentations():

    def __init__(self, application):
        self._application = application

    @property
    def Parent(self):
        return self._application

    def Add(self):
        """Creates a presentation. Returns a Presentation object that represents the new presentation.
        Syntax

        expression.Add (WithWindow)

        expression A variable that represents a Presentations object.
        Parameters
        Name 	Required/Optional 	Data type 	Description
        WithWindow 	Optional 	MsoTriState 	Whether the presentation appears in a visible window.
        Return value

        Presentation
        Example

        pres = Presentations.Add()
        or 
        pres = Application.Presentations.Add()

        """
        return self.Open2007(Path(__file__).parent.parent / "data" /
                             "blank.pptx")

    def Open2007(self, FileName):
        """Opens the specified presentation

        Syntax
        ------
        expression. ``Open2007(FileName)``

        expression A variable that represents a Presentations object.

        Parameter
        ---------

        FileName     Required     String     The name of the file to open.

        Return value
        ------------
        Presentation

        Example
        -------
        ``Presentations.Open2007 "C:\\My Documents\\pres1.ppt"``

        """

        return Presentation.create(FileName)
