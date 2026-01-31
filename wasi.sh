#!/bin/sh

outname=./xlsx2jsonl.wasm
mainpat=./cmd/xlsx2jsonl/main.go

GOOS=wasip1 GOARCH=wasm go \
	build \
	-o "${outname}" \
	-ldflags="-s -w" \
	"${mainpat}"
