package xlsx

import (
	"fmt"
	"math"
	"strconv"
)

type CellType int

const (
	CellTypeString CellType = iota
	CellTypeFormula
	CellTypeNumeric
	CellTypeBool
	CellTypeInline
	CellTypeError
)

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Row      *Row
	Value    string
	formula  string
	style    *Style
	numFmt   string
	date1904 bool
	Hidden   bool
	cellType CellType
	cfmt     *CellFormat
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
	FormattedValue() string
	FormatValue() (string, error)
	ParsedValue() (interface{}, error)
}

func NewCell(r *Row) *Cell {
	return &Cell{style: NewStyle(), Row: r}
}

func (c *Cell) Type() CellType {
	return c.cellType
}

// Set string
func (c *Cell) SetString(s string) {
	c.Value = s
	c.formula = ""
	c.cellType = CellTypeString
}

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.FormattedValue()
}

// Set float
func (c *Cell) SetFloat(n float64) {
	c.SetFloatWithFormat(n, "0.00e+00")
}

/*
	Set float with format. The followings are samples of format samples.

	* "0.00e+00"
	* "0", "#,##0"
	* "0.00", "#,##0.00", "@"
	* "#,##0 ;(#,##0)", "#,##0 ;[red](#,##0)"
	* "#,##0.00;(#,##0.00)", "#,##0.00;[red](#,##0.00)"
	* "0%", "0.00%"
	* "0.00e+00", "##0.0e+0"
*/
func (c *Cell) SetFloatWithFormat(n float64, format string) {
	// tmp value. final value is formatted by FormattedValue() method
	c.Value = fmt.Sprintf("%e", n)
	c.numFmt = format
	c.Value = c.FormattedValue()
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Returns the value of cell as a number
func (c *Cell) Float() (float64, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return math.NaN(), err
	}
	return f, nil
}

// Set a 64-bit integer
func (c *Cell) SetInt64(n int64) {
	c.Value = fmt.Sprintf("%d", n)
	c.numFmt = "0"
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Returns the value of cell as 64-bit integer
func (c *Cell) Int64() (int64, error) {
	f, err := strconv.ParseInt(c.Value, 10, 64)
	if err != nil {
		return -1, err
	}
	return f, nil
}

// Set integer
func (c *Cell) SetInt(n int) {
	c.Value = fmt.Sprintf("%d", n)
	c.numFmt = "0"
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Returns the value of cell as integer
// Has max 53 bits of precision
// See: float64(int64(math.MaxInt))
func (c *Cell) Int() (int, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return -1, err
	}
	return int(f), nil
}

// Set boolean
func (c *Cell) SetBool(b bool) {
	if b {
		c.Value = "1"
	} else {
		c.Value = "0"
	}
	c.cellType = CellTypeBool
}

// Get boolean
func (c *Cell) Bool() bool {
	return c.Value == "1"
}

// Set formula
func (c *Cell) SetFormula(formula string) {
	c.formula = formula
	c.cellType = CellTypeFormula
}

// Returns formula
func (c *Cell) Formula() string {
	return c.formula
}

// GetStyle returns the Style associated with a Cell
func (c *Cell) GetStyle() *Style {
	return c.style
}

// SetStyle sets the style of a cell.
func (c *Cell) SetStyle(style *Style) {
	c.style = style
}

// The number format string is returnable from a cell.
func (c *Cell) GetNumberFormat() string {
	return c.numFmt
}

/*
func (c *Cell) formatToTime(format string) (string, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return "", err
	}
	return TimeFromExcelTime(f, c.date1904).Format(format), nil
}

func (c *Cell) formatToFloat(format string) (string, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return "", err
	}
	return fmt.Sprintf(format, f), nil
}

func (c *Cell) formatToInt(format string) (string, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return "", err
	}
	return fmt.Sprintf(format, int(f)), nil
}
*/

func (c *Cell) Format() *CellFormat {
	if c.cfmt == nil {
		ct := ParseFormat(c.numFmt)
		c.cfmt = &ct
	}
	return c.cfmt
}

func (c *Cell) FormatValue() (FormattedValue, error) {
	return c.Format().FormatValue(c.Value, c.cellType, c.date1904)
}

/*
func (c *Cell) FormatValue() (string, error) {
	if c.cellType == CellTypeError {
		return "", errors.New(c.Value)
	}

	return c.Format().FormatValue(c.Value, c.cellType != CellTypeString, c.date1904)
}
*/

// GoValue converts the cell's text Value to a Go-speicfic type, dependent on the format specified
// for that cell.  It will return either a string, bool, float64, time.Time or time.Duration.
/*
func (c *Cell) GoValue() (fmtType FormatType, fmtSubType FormatSubType, value interface{}, err error) {
	if c.cellType == CellTypeBool {
		// Don't think boolean values can be formatted?
		return BoolFormat, NoSubType, c.Value == "1", nil
	}

	ftype, fsubtype, err := c.Format().FormatType(c.Value, c.cellType != CellTypeString, c.date1904)
	if err != nil {
		return NoType, NoSubType, nil, err
	}
	switch ftype {
	case TextFormat:
		return TextFormat, NoSubType, c.Value, nil

	case NumberFormat:
		fv, _ := strconv.ParseFloat(c.Value, 64)
		return NumberFormat, NoSubType, fv, nil

	case TimeFormat:
		fv, _ := strconv.ParseFloat(c.Value, 64)
		switch fsubtype {
		case Date, Time, DateTime:
			return ftype, fsubtype, TimeFromExcelTime(fv, c.date1904), nil
		case Duration:
			return ftype, fsubtype, DurationFromExcelTime(fv), nil
		}
	}
	panic("Unhandled format")
}
*/

/*

func formatTextTokens(tokens []fmtToken, orgValue string) string {
	var result []byte
	for _, tok := range tokens {
		switch tok.fieldType {
		case tokCellText:
			result = append(result, orgValue...)
		case tokOther:
			result = append(result, tok.other...)
		}
	}
	return string(result)
}

var dateFormats = map[string]string{
	"mm-dd-yy":             "01-02-06",
	"d-mmm-yy":             "2-Jan-06",
	"d-mmm":                "2-Jan",
	"mmm-yy":               "Jan-06",
	"h:mm am/pm":           "3:04 pm",
	"h:mm:ss am/pm":        "3:04:05 pm",
	"h:mm":                 "15:04",
	"h:mm:ss":              "15:04:05",
	"m/d/yy h:mm":          "1/2/06 15:04",
	"mm:ss":                "04:05",
	"yyyy\\-mm\\-dd":       "2006-01-02",
	"mmmm\\-yy":            "January-06",
	"dd/mm/yy":             "02/01/06",
	"hh:mm:ss":             "15:04:05",
	"dd/mm/yy\\ hh:mm":     "02/01/06 15:04",
	"dd/mm/yyyy hh:mm:ss":  "02/01/2006 15:04:05",
	"yy-mm-dd":             "06-01-02",
	"d-mmm-yyyy":           "2-Jan-2006",
	"m/d/yy":               "1/2/06",
	"m/d/yyyy":             "1/2/2006",
	"dd-mmm-yyyy":          "02-Jan-2006",
	"dd/mm/yyyy":           "02/01/2006",
	"mm/dd/yy hh:mm am/pm": "01/02/06 03:04 pm",
	"mm/dd/yyyy hh:mm:ss":  "01/02/2006 15:04:05",
	"yyyy-mm-dd hh:mm:ss":  "2006-01-02 15:04:05",
	// special cases
	"[h]:mm:ss": "",
	"mmss.0":    "",
}

// ParsedValue returns the value as an float64, string or time.Time
func (c *Cell) ParsedValue() (interface{}, error) {
	var numberFormat string = c.GetNumberFormat()
	switch {
	case numberFormat == "general":
		return c.Value, nil

	case numberFormat == "":
		return c.Value, nil

	case numberFormat[len(numberFormat)-1] == '%':
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		return f * 100, nil

	case (numberFormat[0] == '0' || strings.Contains(numberFormat, "#")):
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return nil, err
		}
		return f, nil
	}

	if _, ok := dateFormats[numberFormat]; ok {
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return nil, err
		}
		return TimeFromExcelTime(f, c.date1904), nil
	}

	//return "", fmt.Errorf("Unknown numeric format %q", numberFormat)
	return c.Value, nil
}

*/
// Return the formatted version of the value.
/*
func (c *Cell) FormatValue() (string, error) {
	var numberFormat string = c.GetNumberFormat()
	switch numberFormat {
	case "general":
		return c.Value, nil
	case "":
		return c.Value, nil
	case "0", "#,##0":
		return c.formatToInt("%d")
	case "0.00", "#,##0.00", "@":
		return c.formatToFloat("%.2f")
	case "#,##0 ;(#,##0)", "#,##0 ;[red](#,##0)":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		if f < 0 {
			i := int(math.Abs(f))
			return fmt.Sprintf("(%d)", i), nil
		}
		i := int(f)
		return fmt.Sprintf("%d", i), nil
	case "#,##0.00;(#,##0.00)", "#,##0.00;[red](#,##0.00)":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		if f < 0 {
			return fmt.Sprintf("(%.2f)", f), nil
		}
		return fmt.Sprintf("%.2f", f), nil
	case "0%":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		f = f * 100
		return fmt.Sprintf("%d%%", int(f)), nil
	case "0.00%":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		f = f * 100
		return fmt.Sprintf("%.2f%%", f), nil
	case "0.00e+00", "##0.0e+0":
		return c.formatToFloat("%e")
	case "mm-dd-yy":
		return c.formatToTime("01-02-06")
	case "d-mmm-yy":
		return c.formatToTime("2-Jan-06")
	case "d-mmm":
		return c.formatToTime("2-Jan")
	case "mmm-yy":
		return c.formatToTime("Jan-06")
	case "h:mm am/pm":
		return c.formatToTime("3:04 pm")
	case "h:mm:ss am/pm":
		return c.formatToTime("3:04:05 pm")
	case "h:mm":
		return c.formatToTime("15:04")
	case "h:mm:ss":
		return c.formatToTime("15:04:05")
	case "m/d/yy h:mm":
		return c.formatToTime("1/2/06 15:04")
	case "mm:ss":
		return c.formatToTime("04:05")
	case "[h]:mm:ss":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		t := TimeFromExcelTime(f, c.date1904)
		if t.Hour() > 0 {
			return t.Format("15:04:05"), nil
		}
		return t.Format("04:05"), nil
	case "mmss.0":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return "", err
		}
		t := TimeFromExcelTime(f, c.date1904)
		return fmt.Sprintf("%0d%0d.%d", t.Minute(), t.Second(), t.Nanosecond()/1000), nil

	case "yyyy\\-mm\\-dd":
		return c.formatToTime("2006-01-02")
	case "mmmm\\-yy":
		return c.formatToTime("January-06")
	case "dd/mm/yy":
		return c.formatToTime("02/01/06")
	case "hh:mm:ss":
		return c.formatToTime("15:04:05")
	case "dd/mm/yy\\ hh:mm":
		return c.formatToTime("02/01/06 15:04")
	case "dd/mm/yyyy hh:mm:ss":
		return c.formatToTime("02/01/2006 15:04:05")
	case "yy-mm-dd":
		return c.formatToTime("06-01-02")
	case "d-mmm-yyyy":
		return c.formatToTime("2-Jan-2006")
	case "m/d/yy":
		return c.formatToTime("1/2/06")
	case "m/d/yyyy":
		return c.formatToTime("1/2/2006")
	case "dd-mmm-yyyy":
		return c.formatToTime("02-Jan-2006")
	case "dd/mm/yyyy":
		return c.formatToTime("02/01/2006")
	case "mm/dd/yy hh:mm am/pm":
		return c.formatToTime("01/02/06 03:04 pm")
	case "mm/dd/yyyy hh:mm:ss":
		return c.formatToTime("01/02/2006 15:04:05")
	case "yyyy-mm-dd hh:mm:ss":
		return c.formatToTime("2006-01-02 15:04:05")
	}
	//return "", fmt.Errorf("Unknown numeric format %q", numberFormat)
	return c.Value, nil
}
*/

func (c *Cell) FormattedValue() string {
	v, err := c.FormatValue()
	if err != nil {
		return err.Error()
	}
	return v.FormattedValue
}
