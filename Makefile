# install GNU make to run make
MAKE = make
SPHINX-QUICKSTART = sphinx-quickstart

help:
	@echo "make <OPTIONS>. if using windows, run make in git-bash only"
	@echo OPTIONS:
	@echo "  help"
	@echo "  test"
	@echo "  docclean"
	@echo "  docapidoc"
	@echo "  dochtml"
	@echo "  clean"
	@echo "  setup"
	@echo "  sphinx-quickstart"
	@echo "  build"
	@echo "  testupload"
	@echo "  upload"

.PHONY = help, test, docclean, docapidoc, dochtml, clean, setup-win, build,\
			testupload, upload, tox, sphinx-quickstart

test:
	pytest -s -v  tests/

docs:
	mkdir -p docs

sphinx-quickstart: docs
	. .venv/Scripts/activate; cd docs; sphinx-quickstart
	

docclean:
	$(MAKE) -C docs clean

docapidoc:
	sphinx-apidoc -o ./docs/source -e ./src/opc
	$(MAKE) dochtml

dochtml:
	$(MAKE) docclean
	$(MAKE) -C docs html

clean:
	rm -rf "dist"
	rm -rf .tox

.venv:
	python -m venv .venv

setup: .venv
	. .venv/Scripts/activate; pip install -e .[docs,tests,dists]; pip uninstall -y python-pptx-vba

build:
	$(MAKE) clean
	python -m build

testupload:
	twine upload -r testpypi dist/*

upload:
	twine upload dist/*

tox:
	$(MAKE) clean
	tox
