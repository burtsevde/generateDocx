VENV = $(shell pwd)
VERSION = release_1.0.1

help:
	build - create VERSION directory and .zip file of the version
	clean - delete VERSION directory

.PHONY: build

build:
	mkdir $(VENV)/$(VERSION)/
	cp -r $(VENV)/files_for_release/* $(VENV)/$(VERSION)/
	pyinstaller --noconfirm --onefile --console --name "genDocx.exe" "genDocx.py" --distpath $(VERSION) --clean
	zip -r -D $(VENV)/$(VERSION)/$(VERSION).zip $(VERSION)/* -x "$(VERSION)/*.zip"

clean:
	rm -r $(VENV)/$(VERSION)/
	rm -r $(VENV)/build/

ddoc:
	rm -r $(VENV)/docs/