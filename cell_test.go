package xlsx

import . "gopkg.in/check.v1"

type CellSuite struct{}

var _ = Suite(&CellSuite{})

// Test that we can set and get a Value from a Cell
func (s *CellSuite) TestValueSet(c *C) {
	// Note, this test is fairly pointless, it serves mostly to
	// reinforce that this functionality is important, and should
	// the mechanics of this all change at some point, to remind
	// us not to lose this.
	cell := Cell{}
	cell.Value = "A string"
	c.Assert(cell.Value, Equals, "A string")
}

// Test that GetStyle correctly converts the xlsxStyle.Fonts.
func (s *CellSuite) TestGetStyleWithFonts(c *C) {
	font := NewFont(10, "Calibra")
	style := NewStyle()
	style.Font = *font

	cell := &Cell{Value: "123", style: style}
	style = cell.GetStyle()
	c.Assert(style, NotNil)
	c.Assert(style.Font.Size, Equals, 10)
	c.Assert(style.Font.Name, Equals, "Calibra")
}

// Test that SetStyle correctly translates into a xlsxFont element
func (s *CellSuite) TestSetStyleWithFonts(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	font := NewFont(12, "Calibra")
	style := NewStyle()
	style.Font = *font
	cell.SetStyle(style)
	style = cell.GetStyle()
	xFont, _, _, _, _ := style.makeXLSXStyleElements()
	c.Assert(xFont.Sz.Val, Equals, "12")
	c.Assert(xFont.Name.Val, Equals, "Calibra")
}

// Test that GetStyle correctly converts the xlsxStyle.Fills.
func (s *CellSuite) TestGetStyleWithFills(c *C) {
	fill := *NewFill("solid", "FF000000", "00FF0000")
	style := NewStyle()
	style.Fill = fill
	cell := &Cell{Value: "123", style: style}
	style = cell.GetStyle()
	_, xFill, _, _, _ := style.makeXLSXStyleElements()
	c.Assert(xFill.PatternFill.PatternType, Equals, "solid")
	c.Assert(xFill.PatternFill.BgColor.RGB, Equals, "00FF0000")
	c.Assert(xFill.PatternFill.FgColor.RGB, Equals, "FF000000")
}

// Test that SetStyle correctly updates xlsxStyle.Fills.
func (s *CellSuite) TestSetStyleWithFills(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	fill := NewFill("solid", "00FF0000", "FF000000")
	style := NewStyle()
	style.Fill = *fill
	cell.SetStyle(style)
	style = cell.GetStyle()
	_, xFill, _, _, _ := style.makeXLSXStyleElements()
	xPatternFill := xFill.PatternFill
	c.Assert(xPatternFill.PatternType, Equals, "solid")
	c.Assert(xPatternFill.FgColor.RGB, Equals, "00FF0000")
	c.Assert(xPatternFill.BgColor.RGB, Equals, "FF000000")
}

// Test that GetStyle correctly converts the xlsxStyle.Borders.
func (s *CellSuite) TestGetStyleWithBorders(c *C) {
	border := *NewBorder("thin", "thin", "thin", "thin")
	style := NewStyle()
	style.Border = border
	cell := Cell{Value: "123", style: style}
	style = cell.GetStyle()
	_, _, xBorder, _, _ := style.makeXLSXStyleElements()
	c.Assert(xBorder.Left.Style, Equals, "thin")
	c.Assert(xBorder.Right.Style, Equals, "thin")
	c.Assert(xBorder.Top.Style, Equals, "thin")
	c.Assert(xBorder.Bottom.Style, Equals, "thin")
}

// We can return a string representation of the formatted data
func (l *CellSuite) TestFormattedValue(c *C) {
	cell := Cell{Value: "37947.7500001", cellType: CellTypeNumeric}
	negativeCell := Cell{Value: "-37947.7500001", cellType: CellTypeNumeric}
	smallCell := Cell{Value: "0.007", cellType: CellTypeNumeric}
	earlyCell := Cell{Value: "2.1", cellType: CellTypeNumeric}

	cell.numFmt = "general"
	c.Assert(cell.FormattedValue(), Equals, "37947.7500001")
	negativeCell.numFmt = "general"
	c.Assert(negativeCell.FormattedValue(), Equals, "-37947.7500001")

	cell.numFmt, cell.cfmt = "0", nil

	c.Assert(cell.FormattedValue(), Equals, "37948")

	cell.numFmt, cell.cfmt = "#,##0", nil // For the time being we're not doing
	c.Assert(cell.FormattedValue(), Equals, "37,948")

	cell.numFmt, cell.cfmt = "0.00", nil
	c.Assert(cell.FormattedValue(), Equals, "37947.75")

	cell.numFmt, cell.cfmt = "#,##0.00", nil // For the time being we're not doing
	// this comma formatting, so it'll fall back to the related
	// non-comma form.
	c.Assert(cell.FormattedValue(), Equals, "37,947.75")

	cell.numFmt, cell.cfmt = "#,##0;(#,##0)", nil
	c.Assert(cell.FormattedValue(), Equals, "37,948")
	negativeCell.numFmt, negativeCell.cfmt = "#,##0 ;(#,##0)", nil
	c.Assert(negativeCell.FormattedValue(), Equals, "(37,948)")

	cell.numFmt, cell.cfmt = "#,##0;[red](#,##0)", nil
	c.Assert(cell.FormattedValue(), Equals, "37,948")
	negativeCell.numFmt, negativeCell.cfmt = "#,##0 ;[red](#,##0)", nil
	c.Assert(negativeCell.FormattedValue(), Equals, "(37,948)")

	cell.numFmt, cell.cfmt = "0%", nil
	c.Assert(cell.FormattedValue(), Equals, "3794775%")

	cell.numFmt, cell.cfmt = "0.00%", nil
	c.Assert(cell.FormattedValue(), Equals, "3794775.00%")

	cell.numFmt, cell.cfmt = "0.00E+00", nil
	c.Assert(cell.FormattedValue(), Equals, "3.79E+04")

	cell.numFmt, cell.cfmt = "##0.0e+0", nil
	c.Assert(cell.FormattedValue(), Equals, "3.8E+4")

	cell.numFmt, cell.cfmt = "mm-dd-yy", nil
	c.Assert(cell.FormattedValue(), Equals, "11-22-03")

	cell.numFmt, cell.cfmt = "d-mmm-yy", nil
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-03")
	earlyCell.numFmt, earlyCell.cfmt = "d-mmm-yy", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "2-Jan-00")

	cell.numFmt, cell.cfmt = "d-mmm", nil
	c.Assert(cell.FormattedValue(), Equals, "22-Nov")
	earlyCell.numFmt, earlyCell.cfmt = "d-mmm", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "2-Jan")

	cell.numFmt, cell.cfmt = "mmm-yy", nil
	c.Assert(cell.FormattedValue(), Equals, "Nov-03")

	cell.numFmt, cell.cfmt = "h:mm am/pm", nil
	c.Assert(cell.FormattedValue(), Equals, "6:00 pm")
	smallCell.numFmt, smallCell.cfmt = "h:mm am/pm", nil
	c.Assert(smallCell.FormattedValue(), Equals, "12:10 am")

	cell.numFmt, cell.cfmt = "h:mm:ss am/pm", nil
	c.Assert(cell.FormattedValue(), Equals, "6:00:00 pm")
	smallCell.numFmt, smallCell.cfmt = "h:mm:ss am/pm", nil
	c.Assert(smallCell.FormattedValue(), Equals, "12:10:05 am")

	cell.numFmt, cell.cfmt = "h:mm", nil
	c.Assert(cell.FormattedValue(), Equals, "18:00")
	smallCell.numFmt, smallCell.cfmt = "h:mm", nil
	c.Assert(smallCell.FormattedValue(), Equals, "00:10")

	cell.numFmt, cell.cfmt = "h:mm:ss", nil
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	// This is wrong, but there's no eary way aroud it in Go right now, AFAICT.
	smallCell.numFmt, smallCell.cfmt = "h:mm:ss", nil
	c.Assert(smallCell.FormattedValue(), Equals, "00:10:05")

	cell.numFmt, cell.cfmt = "m/d/yy h:mm", nil
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 18:00")
	smallCell.numFmt, smallCell.cfmt = "m/d/yy h:mm", nil
	c.Assert(smallCell.FormattedValue(), Equals, "12/31/99 00:10") // Note, that's 1899
	earlyCell.numFmt, earlyCell.cfmt = "m/d/yy h:mm", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "1/2/00 02:24") // and 1900

	cell.numFmt, cell.cfmt = "mm:ss", nil
	c.Assert(cell.FormattedValue(), Equals, "00:00")
	smallCell.numFmt, smallCell.cfmt = "mm:ss", nil
	c.Assert(smallCell.FormattedValue(), Equals, "10:05")

	earlyCell.numFmt, earlyCell.cfmt = "[h]:mm:ss", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "50:24:00")

	cell.numFmt, cell.cfmt = "mmss.0", nil // I'm not sure about these.
	//c.Assert(cell.FormattedValue(), Equals, "00.8640")
	smallCell.numFmt, smallCell.cfmt = "mmss.0", nil
	c.Assert(smallCell.FormattedValue(), Equals, "1004.8")
	//c.Assert(smallCell.FormattedValue(), Equals, "1447.999997")

	cell.numFmt, cell.cfmt = "yyyy\\-mm\\-dd", nil
	c.Assert(cell.FormattedValue(), Equals, "2003-11-22")

	cell.numFmt, cell.cfmt = "dd/mm/yy", nil
	c.Assert(cell.FormattedValue(), Equals, "22/11/03")
	earlyCell.numFmt, earlyCell.cfmt = "dd/mm/yy", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "02/01/00")

	cell.numFmt, cell.cfmt = "hh:mm:ss", nil
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	smallCell.numFmt, smallCell.cfmt = "hh:mm:ss", nil
	c.Assert(smallCell.FormattedValue(), Equals, "00:10:05")

	cell.numFmt, cell.cfmt = "dd/mm/yy\\ hh:mm", nil
	c.Assert(cell.FormattedValue(), Equals, "22/11/03 18:00")

	cell.numFmt, cell.cfmt = "yy-mm-dd", nil
	c.Assert(cell.FormattedValue(), Equals, "03-11-22")

	cell.numFmt, cell.cfmt = "d-mmm-yyyy", nil
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")
	earlyCell.numFmt, earlyCell.cfmt = "d-mmm-yyyy", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "2-Jan-1900")

	cell.numFmt, cell.cfmt = "m/d/yy", nil
	c.Assert(cell.FormattedValue(), Equals, "11/22/03")
	earlyCell.numFmt, earlyCell.cfmt = "m/d/yy", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "1/2/00")

	cell.numFmt, cell.cfmt = "m/d/yyyy", nil
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003")
	earlyCell.numFmt, earlyCell.cfmt = "m/d/yyyy", nil
	c.Assert(earlyCell.FormattedValue(), Equals, "1/2/1900")

	cell.numFmt, cell.cfmt = "dd-mmm-yyyy", nil
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")

	cell.numFmt, cell.cfmt = "dd/mm/yyyy", nil
	c.Assert(cell.FormattedValue(), Equals, "22/11/2003")

	cell.numFmt, cell.cfmt = "mm/dd/yy hh:mm am/pm", nil
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 06:00 pm")

	cell.numFmt, cell.cfmt = "mm/dd/yyyy hh:mm:ss", nil
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003 18:00:00")
	smallCell.numFmt, smallCell.cfmt = "mm/dd/yyyy hh:mm:ss", nil
	c.Assert(smallCell.FormattedValue(), Equals, "12/31/1899 00:10:05")

	cell.numFmt, cell.cfmt = "yyyy-mm-dd hh:mm:ss", nil
	c.Assert(cell.FormattedValue(), Equals, "2003-11-22 18:00:00")
	smallCell.numFmt, smallCell.cfmt = "yyyy-mm-dd hh:mm:ss", nil
	c.Assert(smallCell.FormattedValue(), Equals, "1899-12-31 00:10:05")
}

// test setters and getters
func (s *CellSuite) TestSetterGetters(c *C) {
	cell := Cell{}

	cell.SetString("hello world")
	c.Assert(cell.String(), Equals, "hello world")
	c.Assert(cell.Type(), Equals, CellTypeString)

	cell.SetInt(1024)
	intValue, _ := cell.Int()
	c.Assert(intValue, Equals, 1024)
	c.Assert(cell.Type(), Equals, CellTypeNumeric)

	cell.SetFloat(1.024)
	float, _ := cell.Float()
	intValue, _ = cell.Int() // convert
	c.Assert(float, Equals, 1.024)
	c.Assert(intValue, Equals, 1)
	c.Assert(cell.Type(), Equals, CellTypeNumeric)

	cell.SetFormula("10+20")
	c.Assert(cell.Formula(), Equals, "10+20")
	c.Assert(cell.Type(), Equals, CellTypeFormula)
}

/*
var tt = time.Date(2015, 4, 5, 15, 0, 0, 0, time.UTC)
var goValueTests = []struct {
	value           string
	cellType        CellType
	fmt             string
	expectedType    FormatType
	expectedSubType FormatSubType
	expectedGoValue interface{}
}{
	{"foo", CellTypeString, "", TextFormat, NoSubType, "foo"},
	{"1", CellTypeString, "", TextFormat, NoSubType, "1"},
	{"1", CellTypeNumeric, "", NumberFormat, NoSubType, float64(1)},
	{"42099.625", CellTypeNumeric, "yyyy-mm-dd", TimeFormat, Date, tt},
	{"42099.625", CellTypeNumeric, "hh:mm", TimeFormat, Time, tt},
	{"42099.625", CellTypeNumeric, "yyyy-mm-dd hh:mm", TimeFormat, DateTime, tt},
	{"2.5", CellTypeNumeric, "[hh]:mm", TimeFormat, Duration, 60 * time.Hour},
	{"-2.5", CellTypeNumeric, "[hh]:mm;##", NumberFormat, NoSubType, -2.5},
}

func (s *CellSuite) TestGoValue(c *C) {
	for _, test := range goValueTests {
		cell := Cell{Value: test.value, numFmt: test.fmt, cellType: test.cellType}
		ftype, fsubtype, v, err := cell.GoValue()
		c.Assert(err, IsNil)
		c.Assert(ftype, Equals, test.expectedType, Commentf("value=%q", test.value))
		c.Assert(fsubtype, Equals, test.expectedSubType, Commentf("value=%q", test.value))
		c.Assert(v, Equals, test.expectedGoValue, Commentf("value=%q", test.value))
	}
}
*/
