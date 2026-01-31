#!/bin/sh

outname=./xlsx2jsonl.wasm
mname=./cmd/xlsx2jsonl/main.go

tinygo \
	build \
	-o "${outname}" \
	-target=wasip1 \
	-opt=z \
	-no-debug \
	"${mname}"
