DEPS=$(wildcard bits/*.js)
TARGET=xls.js

$(TARGET): $(DEPS)
	cat $^ > $@

bits/01_version.js: package.json
	echo "XLS.version = '"`grep version package.json | awk '{gsub(/[^0-9\.]/,"",$$2); print $$2}'`"';" > bits/01_version.js

.PHONY: clean
clean:
	rm $(TARGET)

.PHONY: init
init:
	git submodule init
	git submodule update
	git submodule foreach git pull origin master
	git submodule foreach make

.PHONY: oldtest
oldtest:
	bin/test.sh

.PHONY: test mocha
test mocha: test.js
	mocha -R spec

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)

.PHONY: cov
cov: misc/coverage.html

misc/coverage.html: xls.js test.js
	mocha --require blanket -R html-cov > misc/coverage.html

.PHONY: coveralls
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js
