.PHONY: build install clean test check

BINARY   = outlook-ax
PREFIX  ?= /usr/local/bin

build:
	swift build -c release --product $(BINARY)
	cp .build/release/$(BINARY) .

install: build
	cp $(BINARY) $(PREFIX)/$(BINARY)

test: build
	swift test

clean:
	rm -f $(BINARY)
	swift package clean

check: build test
	bash scripts/check-repository-layout.sh
	bash scripts/check-architecture.sh
