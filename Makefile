DEPS=$(wildcard bits/*.js)
TARGET=xls.js

$(TARGET): $(DEPS)
	cat $^ > $@

.PHONY: clean
clean:
	rm $(TARGET)

.PHONY: init
init:
	git submodule init
	git submodule update
	git submodule foreach git pull origin master 

.PHONY: oldtest
oldtest:
	bin/test.sh

.PHONY: test mocha
test mocha:
	mocha -R spec

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)
