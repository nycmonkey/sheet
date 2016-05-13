package sheet

import (
	"errors"
	"io"
	"io/ioutil"

	"github.com/extrame/xls"
)

type xlsReader struct{}

func (xr *xlsReader) Read(r io.Reader, s Specification) (si *Info, err error) {
	var wb *xls.WorkBook
	wb, err = xls.OpenReader(ioutil.NopCloser(r), "CP-1252") // charset parameter isn't used in the underlying library
	if err != nil {
		return
	}
	var ws *xls.WorkSheet
	if s.SheetName != "" {
		found := false
		for i := 0; i < wb.NumSheets(); i++ {
			ws = wb.GetSheet(i)
			if ws == nil {
				continue
			}
			if ws.Name == s.SheetName {
				found = true
				break
			}
		}
		if !found {
			err = errors.New("Workbook does not include a sheet named " + s.SheetName)
			return
		}
	}
	if ws == nil {
		ws = wb.GetSheet(s.SheetIndex)
	}
	if ws == nil {
		err = errors.New("Failed to open the worksheet")
		return
	}
	si = &Info{}
	for k, row := range ws.Rows {
		if k == s.HeaderRowIndex {
			si.ColCount = len(row.Cols)
			continue
		}
		if k >= s.DataRowStartIndex {
			si.RowCount++
		}
	}
	return
}
