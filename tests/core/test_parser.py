import pytest
from ppv.core.parser import ElementBase


def test_parsed_xml_is_of_elementbase_type(pres_xml):
    assert isinstance(pres_xml, ElementBase)


def test_elementbase_findqn(pres_xml):
    assert pres_xml.findqn('p:sldIdLst').tag == "{"+pres_xml.ns+"}"+"sldIdLst"


def test_elementbase_findallqn(pres_xml):
    e = pres_xml.findqn('p:sldIdLst')
    assert len(e.findallqn('p:sldId')) == 2


def test_elementbase_getqn(pres_xml):
    e = pres_xml.findqn('p:sldIdLst').findqn('p:sldId')
    assert e.getqn('id') == '256'


def test_elementbase_getqn_default(pres_xml):
    e = pres_xml.findqn('p:sldIdLst').findqn('p:sldId')
    assert e.getqn('NA', -1) == -1
