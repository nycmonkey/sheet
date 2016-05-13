package sheet

import "testing"

func TestXlsReader(t *testing.T) {
	spec1 := Specification{
		SheetName:         "Data",
		HeaderRowIndex:    2,
		DataRowStartIndex: 3,
	}
	i, err := ParseFile("starts on row 3.xls", spec1)
	if err != nil {
		t.Errorf("Did not expect an error, but got %s", err)
	}
	expectedCount := 3
	if i.ColCount != expectedCount {
		t.Errorf("Expected column count of %d, got %d", expectedCount, i.ColCount)
	}
	if i.RowCount != expectedCount {
		t.Errorf("Expected row count of %d, got %d", expectedCount, i.RowCount)
	}
}

func TestGfdFile(t *testing.T) {
	spec1 := Specification{
		HeaderRowIndex:    5,
		DataRowStartIndex: 6,
	}
	i, err := ParseFile(`C:\dev\etl\Relationship Data Source Files\gfdGfiaGeq.xls.xls`, spec1)
	if err != nil {
		t.Errorf("Did not expect an error, but got %s", err)
	}
	expectedColCount := 12
	if i.ColCount != expectedColCount {
		t.Errorf("Expected column count of %d, got %d", expectedColCount, i.ColCount)
	}
	expectedRowCount := 23091
	if i.RowCount != expectedRowCount {
		t.Errorf("Expected row count of %d, got %d", expectedRowCount, i.RowCount)
	}
}
