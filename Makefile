#!/usr/bin/make -f

CFLAGS = -std=c++11

CFLAGS += $(shell pkg-config --cflags libpng)
LDFLAGS += $(shell pkg-config --libs libpng)

all: lok_convert lok_split lok_keyevent lok_screenshot

lok_convert: lok_convert.cpp
	g++ $^ -Wall -Werror -ldl -o $@

lok_split: lok_split.cpp
	g++ $^ -Wall -Werror ${CFLAGS} ${LDFLAGS} -ldl -o $@

lok_keyevent: lok_keyevent.cpp
	g++ $^ -Wall -Werror -std=c++11 -ldl -o $@

lok_screenshot: lok_screenshot.cpp
	g++ $^ -Wall -Werror -std=c++11 -ldl -o $@
