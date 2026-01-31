#!/bin/sh

go \
	vet \
	-all \
	-race \
	./... || exec sh -c 'echo go vet failure.; exit 1'

golangci-lint \
	run \
	--config ./.golangci.yml

find . -type f -name '*.go' |
	xargs \
		gopls \
		check

find . -type f -name '*.go' | xargs fgrep '//nolint' | fgrep depguard && exec sh -c '
	echo nolint for depguard forbidden.
	exit 1
'
