.PHONY: build install clean

BINARY = outlook-ax
SRC = outlook-ax.swift
PREFIX ?= /usr/local/bin

build:
	swiftc -O $(SRC) -o $(BINARY)

install: build
	cp $(BINARY) $(PREFIX)/$(BINARY)

clean:
	rm -f $(BINARY)
