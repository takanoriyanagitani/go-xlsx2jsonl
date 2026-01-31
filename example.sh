#!/bin/sh

ixlsx="./sample.d/input.xlsx"
sname=Sheet4

geninput(){
	wasmloc="${HOME}/.cargo/bin/rs-jsonl2xlsx.wasm"
	test -f "${wasmloc}" || exec sh -c '
		echo rs-jsonl2xlsx.wasm not installed.
		echo see github.com/takanoriyanagitani/rs-jsonl2xlsx to install it.
		exit 1
	'

	echo generating the input xlsx file...

	mkdir -p ./sample.d

	jq -c -n '[
	  {id:42, date:"2026-02-01", severity:"INFO", status:200, body:"apt update done"},
	  {id:43, date:"2026-02-02", severity:"WARN", status:500, body:"apt update failure"}
	]' |
		jq -c '.[]' |
		wazero \
			run "${wasmloc}" \
			-- \
			--sheet-name "${sname}" |
		dd if=/dev/stdin of="${ixlsx}" bs=1048576 status=none
}

run_wasi(){
	cat "${ixlsx}" |
		wasmtime \
			run \
			./xlsx2jsonl.wasm \
			-sheet-name "${sname}" \
			-skip-rows 0 |
		jq -c
}

test -s "${ixlsx}" || geninput

test -f "${ixlsx}" || exec env ix="${ixlsx}" sh -c '
	echo xlsx file "${ix}" missing.
	exit 1
'

test -s "${ixlsx}" || exec env ix="${ixlsx}" sh -c '
	echo the file "${ix}" is empty.
	exit 1
'

run_wasi
