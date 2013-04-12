DEPS=$(wildcard bits/*.js)
TARGET=xls.js
$(TARGET): $(DEPS)
	cat $^ > $@

.PHONY: clean
clean:
	rm $(TARGET) 
