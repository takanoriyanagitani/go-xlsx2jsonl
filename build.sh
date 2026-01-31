#!/bin/sh

bname="xlsx2jsonl"
bdir="./cmd/${bname}"
oname="${bdir}/${bname}"

go \
	build \
	-v \
	./...

go \
	build \
	-v \
	-o "${oname}" \
	"${bdir}"
