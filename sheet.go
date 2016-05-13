package sheet

import (
	"errors"
	"io"
	"os"
	"path/filepath"
)

// Info holds the heading row and row count parsed by the reader
type Info struct {
	ColCount int
	RowCount int
}

// Specification configures a SheetReader for the nuances of a particular source
type Specification struct {
	SheetName         string
	SheetIndex        int
	HeaderRowIndex    uint16
	DataRowStartIndex uint16
}

func readerFor(ext string) (Reader, error) {
	switch ext {
	case ".xls":
		return &xlsReader{}, nil
	}
	return nil, errors.New("No parser available for extension " + ext)
}

// ParseFile returns Info for a spreadsheet
func ParseFile(path string, s Specification) (i *Info, err error) {
	var parser Reader
	parser, err = readerFor(filepath.Ext(path))
	var f *os.File
	f, err = os.Open(path)
	if err != nil {
		return
	}
	defer f.Close()
	return parser.Read(f, s)
}

// Reader is implemented by parsers that can read a spreadsheet
type Reader interface {
	Read(r io.Reader, s Specification) (si *Info, err error)
}
