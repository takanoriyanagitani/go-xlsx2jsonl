package x2jsonl

import (
	"bufio"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"iter"
	"log"
	"os"
	"slices"
	"strconv"

	xpkg "github.com/xuri/excelize/v2"
)

var (
	ErrUnableToGetHeader  error = errors.New("unable to get header")
	ErrInvalidColumnCount error = errors.New("invalid column count")
	ErrNoSampleRow        error = errors.New("too few rows")
	ErrTooManyColumns     error = errors.New("too many columns")
	ErrUnexpectedTypCount error = errors.New("column count mismatch")
)

type Row struct {
	Index   int
	Columns []string
}

type Xfile struct{ *xpkg.File }

type SheetName = string
type CellAddress = string

func (x Xfile) Close() error { return x.File.Close() }

type Xtype byte

type RowIndex uint32
type ColIndex uint16

const (
	XtypeUnset        Xtype = Xtype(xpkg.CellTypeUnset)
	XtypeBool         Xtype = Xtype(xpkg.CellTypeBool)
	XtypeDate         Xtype = Xtype(xpkg.CellTypeDate)
	XtypeError        Xtype = Xtype(xpkg.CellTypeError)
	XtypeFormula      Xtype = Xtype(xpkg.CellTypeFormula)
	XtypeInlineString Xtype = Xtype(xpkg.CellTypeInlineString)
	XtypeNumber       Xtype = Xtype(xpkg.CellTypeNumber)
	XtypeSharedString Xtype = Xtype(xpkg.CellTypeSharedString)
)

func (x Xfile) CellType(sheet SheetName, addr CellAddress) (Xtype, error) {
	typ, e := x.File.GetCellType(sheet, addr)
	if nil != e {
		return XtypeUnset, e
	}

	return Xtype(typ), nil
}

func (x Xfile) CellTypeByIndex(
	sheet SheetName,
	row RowIndex,
	col ColIndex,
) (Xtype, error) {
	addr, err := xpkg.CoordinatesToCellName(
		int(col),
		int(row),
		false,
	)
	if nil != err {
		return XtypeUnset, err
	}

	return x.CellType(sheet, addr)
}

func (x Xfile) ToRows(sheet SheetName, skipRows int) iter.Seq2[Row, error] {
	var rows *xpkg.Rows
	var e error
	rows, e = x.File.Rows(sheet)
	if nil != e {
		return func(yield func(Row, error) bool) {
			yield(Row{}, e)
		}
	}

	return Xrows{Rows: rows}.ToIter(skipRows)
}

func (x Xfile) ToHeader(sheet SheetName, skipRows int) ([]string, error) {
	var rowsExcel *xpkg.Rows
	var e error
	rowsExcel, e = x.File.Rows(sheet)
	if nil != e {
		return nil, e
	}
	var rowsIter iter.Seq2[Row, error] = Xrows{Rows: rowsExcel}.ToIter(skipRows)

	for row, e := range rowsIter {
		if nil != e {
			return nil, e
		}
		return row.Columns, nil
	}
	return nil, fmt.Errorf("%w: sheet=%s", ErrUnableToGetHeader, sheet)
}

func (x Xfile) ToTypedValue(sheet SheetName, rowpos uint32, values []string) ([]TypedValue, error) {
	ret := make([]TypedValue, 0, len(values))
	for colix, col := range values {
		if 65535 < colix { //nolint:mnd
			return nil, ErrTooManyColumns
		}
		var colpos uint16 = uint16(colix) + 1 //nolint:gosec
		xtyp, e := x.CellTypeByIndex(sheet, RowIndex(rowpos), ColIndex(colpos))
		if nil != e {
			return nil, e
		}

		ret = append(ret, TypedValue{
			Raw: col,
			Typ: xtyp,
		})
	}
	return ret, nil
}

func (x Xfile) ToSampleRow(sheet SheetName, skipRows int) ([]TypedValue, error) {
	var rowsExcel *xpkg.Rows
	var e error
	rowsExcel, e = x.File.Rows(sheet)
	if nil != e {
		return nil, e
	}
	var rowsIter iter.Seq2[Row, error] = Xrows{Rows: rowsExcel}.ToIter(skipRows)

	var rowCount int = 0
	for row, e := range rowsIter {
		if nil != e {
			return nil, e
		}
		rowCount++
		if rowCount == 1 { // This is the header row, we need the next one for sample.
			continue
		}
		physicalRowIndex := uint32(row.Index) //nolint:gosec
		return x.ToTypedValue(sheet, physicalRowIndex, row.Columns)
	}
	return nil, ErrNoSampleRow
}

func (x Xfile) SheetNames() []string {
	return x.File.GetSheetList()
}

func (x Xfile) ToRawObjects(sheet SheetName, skipRows int) iter.Seq2[map[string]string, error] {
	hdrs, err := x.ToHeader(sheet, skipRows)
	if nil != err {
		return func(yield func(map[string]string, error) bool) {
			yield(nil, err)
		}
	}

	rows, err := x.File.Rows(sheet)
	if nil != err {
		return func(yield func(map[string]string, error) bool) {
			yield(nil, err)
		}
	}

	return Xrows{Rows: rows}.ToRawObjects(hdrs, skipRows)
}

func (x Xfile) ToObjects(sheet SheetName, skipRows int) iter.Seq2[map[string]any, error] {
	hdrs, err := x.ToHeader(sheet, skipRows)
	if nil != err {
		return func(yield func(map[string]any, error) bool) {
			yield(nil, err)
		}
	}

	rows, err := x.File.Rows(sheet)
	if nil != err {
		return func(yield func(map[string]any, error) bool) {
			yield(nil, err)
		}
	}

	typed, err := x.ToSampleRow(sheet, skipRows)
	if nil != err {
		return func(yield func(map[string]any, error) bool) {
			yield(nil, err)
		}
	}

	var ityps iter.Seq[Xtype] = func(yield func(Xtype) bool) {
		for _, typedVal := range typed {
			if !yield(typedVal.Typ) {
				return
			}
		}
	}

	var typs []Xtype = slices.Collect(ityps)

	return Xrows{Rows: rows}.ToObjects(hdrs, typs, skipRows)
}

func (x Xfile) RawsToJsonsToWriter(sheet SheetName, wtr io.Writer, skipRows int) error {
	var bw *bufio.Writer = bufio.NewWriter(wtr)
	var raws iter.Seq2[map[string]string, error] = x.ToRawObjects(sheet, skipRows)
	var enc *json.Encoder = json.NewEncoder(bw)
	e := JsonEnc{Encoder: enc}.WriteRawObjects(raws)
	return errors.Join(e, bw.Flush())
}

func (x Xfile) ToJsonsToWriter(sheet SheetName, wtr io.Writer, skipRows int) error {
	var bw *bufio.Writer = bufio.NewWriter(wtr)
	var objs iter.Seq2[map[string]any, error] = x.ToObjects(sheet, skipRows)
	var enc *json.Encoder = json.NewEncoder(bw)
	e := JsonEnc{Encoder: enc}.WriteObjects(objs)
	return errors.Join(e, bw.Flush())
}

func (x Xfile) RawsToJsonsToStdout(sheet SheetName, skipRows int) error {
	return x.RawsToJsonsToWriter(sheet, os.Stdout, skipRows)
}

func (x Xfile) ObjsToJsonsToStdout(sheet SheetName, skipRows int) error {
	return x.ToJsonsToWriter(sheet, os.Stdout, skipRows)
}

func ReaderToXfile(rdr io.Reader) (Xfile, error) {
	f, e := xpkg.OpenReader(rdr)
	return Xfile{File: f}, e
}

func StdinToSheetToRawJsonsToStdout(sheet SheetName, skipRows int) error {
	xfile, e := ReaderToXfile(os.Stdin)
	if nil != e {
		return e
	}

	defer func() {
		e := xfile.Close()
		if nil != e {
			log.Printf("close error: %v\n", e)
		}
	}()

	return xfile.RawsToJsonsToStdout(sheet, skipRows)
}

func StdinToSheetToJsonsToStdout(sheet SheetName, skipRows int) error {
	xfile, e := ReaderToXfile(os.Stdin)
	if nil != e {
		return e
	}

	defer func() {
		e := xfile.Close()
		if nil != e {
			log.Printf("close error: %v\n", e)
		}
	}()

	return xfile.ObjsToJsonsToStdout(sheet, skipRows)
}

type Xrows struct{ *xpkg.Rows }

func (r Xrows) Close() error { return r.Rows.Close() }

func (r Xrows) ToIter(skipRows int) iter.Seq2[Row, error] {
	return func(yield func(Row, error) bool) {
		defer func() {
			e := r.Close()
			if nil != e {
				log.Printf("error on row close: %v\n", e)
			}
		}()

		var currentPhysicalRow int = 1
		for currentPhysicalRow <= skipRows {
			if !r.Rows.Next() {
				// If we exhaust rows before skipping enough, yield no error, just finish.
				return
			}
			err := r.Rows.Error()
			if nil != err {
				yield(Row{}, err) // Yield error if encountered during skipping
				return
			}
			currentPhysicalRow++
		}

		for r.Rows.Next() {
			err := r.Rows.Error()
			if nil != err {
				yield(Row{}, err)
				return
			}

			cols, err := r.Rows.Columns()
			if !yield(Row{Index: currentPhysicalRow, Columns: cols}, err) {
				return
			}
			currentPhysicalRow++
		}
	}
}

func (r Xrows) ToRawObjects(headers []string, skipRows int) iter.Seq2[map[string]string, error] {
	return func(yield func(map[string]string, error) bool) {
		var rows iter.Seq2[Row, error] = r.ToIter(skipRows)

		buf := map[string]string{}

		var firstRowSkipped bool = false // To skip the actual header row (which is the first yielded after skipRows)

		for row, e := range rows {
			clear(buf)
			if nil != e {
				yield(nil, e)
				return
			}

			if !firstRowSkipped {
				firstRowSkipped = true
				continue // Skip the header row itself
			}

			cols := row.Columns

			var colcnt int = len(cols)
			var hdrcnt int = len(headers)
			if colcnt != hdrcnt {
				yield(
					nil,
					fmt.Errorf(
						"%w: column count=%v, header count=%v",
						ErrInvalidColumnCount,
						colcnt,
						hdrcnt,
					),
				)
				return
			}

			for i := range colcnt {
				var hdr string = headers[i]
				var col string = cols[i]
				buf[hdr] = col
			}

			if !yield(buf, nil) {
				return
			}
		}
	}
}


//nolint:funlen
func (r Xrows) ToObjects(headers []string, typs []Xtype, skipRows int) iter.Seq2[map[string]any, error] {
	if len(headers) != len(typs) {
		return func(yield func(map[string]any, error) bool) {
			yield(nil, ErrUnexpectedTypCount)
		}
	}

	return func(yield func(map[string]any, error) bool) {
		var rows iter.Seq2[Row, error] = r.ToIter(skipRows)

		buf := map[string]any{}

		var firstRowSkipped bool = false // To skip the actual header row (which is the first yielded after skipRows)

		for row, e := range rows {
			clear(buf)
			if nil != e {
				yield(nil, e)
				return
			}

			if !firstRowSkipped {
				firstRowSkipped = true
				continue // Skip the header row itself
			}

			cols := row.Columns

			var colcnt int = len(cols)
			var hdrcnt int = len(headers)
			if colcnt != hdrcnt {
				yield(
					nil,
					fmt.Errorf(
						"%w: column count=%v, header count=%v",
						ErrInvalidColumnCount,
						colcnt,
						hdrcnt,
					),
				)
				return
			}

			for i := range colcnt {
				var hdr string = headers[i]
				var col string = cols[i]
				var typ Xtype = typs[i]
				typed := TypedValue{
					Raw: col,
					Typ: typ,
				}
				converted, e := typed.Convert()
				if nil != e {
					yield(nil, e)
					return
				}
				buf[hdr] = converted
			}

			if !yield(buf, nil) {
				return
			}
		}
	}
}

type TypedValue struct {
	Raw string
	Typ Xtype
}

func (v TypedValue) ConvertUnset() (any, error) { return nil, nil } //nolint:nilnil

func (v TypedValue) ConvertBool() (bool, error)      { return strconv.ParseBool(v.Raw) }
func (v TypedValue) ConvertNumber() (float64, error) { return strconv.ParseFloat(v.Raw, 64) }
func (v TypedValue) ConvertString() (string, error)  { return v.Raw, nil }

func (v TypedValue) TryConvertUnset() (any, error) {
	if "" == v.Raw {
		return v.ConvertUnset()
	}

	f, err := strconv.ParseFloat(v.Raw, 64)
	if nil == err {
		return f, nil
	}

	b, err := strconv.ParseBool(v.Raw)
	if nil == err {
		return b, nil
	}

	i, err := strconv.ParseInt(v.Raw, 10, 64)
	if nil == err {
		return i, nil
	}

	return v.Raw, nil
}

func (v TypedValue) Convert() (any, error) {
	switch v.Typ {
	case XtypeUnset:
		return v.TryConvertUnset()
	case XtypeBool:
		return v.ConvertBool()
	case XtypeDate:
		return v.ConvertString()
	case XtypeError:
		return v.ConvertString()
	case XtypeFormula:
		return v.ConvertString()
	case XtypeInlineString:
		return v.ConvertString()
	case XtypeNumber:
		return v.ConvertNumber()
	case XtypeSharedString:
		return v.ConvertString()
	default:
		return v.ConvertString()
	}
}

type JsonEnc struct{ *json.Encoder }

func (j JsonEnc) WriteRawObjects(raws iter.Seq2[map[string]string, error]) error {
	for raw, e := range raws {
		if nil != e {
			return e
		}

		je := j.Encoder.Encode(raw)
		if nil != je {
			return je
		}
	}
	return nil
}

func (j JsonEnc) WriteObjects(objs iter.Seq2[map[string]any, error]) error {
	for obj, e := range objs {
		if nil != e {
			return e
		}

		je := j.Encoder.Encode(obj)
		if nil != je {
			return je
		}
	}
	return nil
}
