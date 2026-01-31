package main

import (
	"flag"
	"log"

	xj "github.com/takanoriyanagitani/go-xlsx2jsonl"
)

func main() {
	var sname string
	flag.StringVar(&sname, "sheet-name", "Sheet1", "sheet name to convert")
	var sr int
	flag.IntVar(&sr, "skip-rows", 0, "skip rows")
	flag.Parse()

	e := xj.StdinToSheetToJsonsToStdout(sname, sr)
	if nil != e {
		log.Printf("%v\n", e)
	}
}
