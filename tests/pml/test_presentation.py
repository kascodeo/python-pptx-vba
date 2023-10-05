def test_Parent(blank_pres, application):
    assert application.Presentations is blank_pres.Parent


def test_SaveAs(blank_pres, saved_path):
    blank_pres.SaveAs(saved_path)


def test_BuiltInDocumentProperties(blank_pres):
    from opc.coreprops import CoreProperties
    assert isinstance(blank_pres.BuiltInDocumentProperties, CoreProperties)


def test_HandoutMaster(allinone_pres):
    from ppv.pml.handoutmaster import HandoutMaster
    assert isinstance(allinone_pres.HandoutMaster, HandoutMaster)


def test_HasHandoutMaster(allinone_pres):
    assert allinone_pres.HasHandoutMaster


def test_HasNotesMaster(allinone_pres):
    assert allinone_pres.HasNotesMaster


def test_NotesMaster(allinone_pres):
    from ppv.pml.notesmaster import NotesMaster
    assert isinstance(allinone_pres.NotesMaster, NotesMaster)


def test_SlideMaster(blank_pres):
    from ppv.pml.slidemaster import SlideMaster
    assert isinstance(blank_pres.SlideMaster, SlideMaster)


def test_Slides(blank_pres):
    from ppv.pml.slides import Slides
    assert isinstance(blank_pres.Slides, Slides)
