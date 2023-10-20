import pytest
from pathlib import Path
from ppv import Application
from ppv.core.parser import Parser
from ppv.enum.MsoTextOrientation import MsoTextOrientation


@pytest.fixture
def application():
    return Application


@pytest.fixture
def presentations():
    return Application.Presentations


@pytest.fixture
def data_path():
    return Path(__file__).parent / "data"


@pytest.fixture
def blank_path(data_path):
    return data_path / "blank.pptx"


@pytest.fixture
def shapes_path(data_path):
    return data_path / "shapes.pptx"


@pytest.fixture
def allinone_path(data_path):
    return data_path / "allinone.pptx"


@pytest.fixture
def saved_path(data_path):
    return data_path / "saved.pptx"


@pytest.fixture
def blank_pres(blank_path):
    return Application.Presentations.Open2007(blank_path)


@pytest.fixture
def shapes_pres(shapes_path):
    return Application.Presentations.Open2007(shapes_path)


@pytest.fixture
def slide(blank_pres):
    return blank_pres.Slides.Item(1)


@pytest.fixture
def slide_layout(blank_pres):
    pkg = blank_pres._pkg
    return pkg.get_part('/ppt/slideLayouts/slideLayout7.xml').typeobj


@pytest.fixture
def allinone_pres(allinone_path):
    return Application.Presentations.Open2007(allinone_path)


@pytest.fixture
def pres_xml_path(data_path):
    return data_path / "presentation.xml"


@pytest.fixture
def parser():
    return Parser()


@pytest.fixture
def pres_xml(parser, pres_xml_path):
    with open(pres_xml_path) as f:
        return parser.parse(f).getroot()


@pytest.fixture
def shapes(shapes_pres):
    return shapes_pres.Slides.Item(1).Shapes


@pytest.fixture
def shape_of_table(shapes):
    return shapes.Item(6)


@pytest.fixture
def textbox(shapes):
    return shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                             10, 10, 200, 10)


@pytest.fixture
def table(shape_of_table):
    return shape_of_table.Table
