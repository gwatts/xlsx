package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/kr/pretty"
)

var (
	UnsupportedCondition = errors.New("Unsupported Condition")
)

type FmtTokenType int

const (
	TokInvalid  FmtTokenType = iota
	TokCellText              // "@" symbol to be replaced with cell contents
	TokGeneral               // Apply general numeric formatting
	TokColor

	TokNumInt
	TokNumDec
	TokNumDecSign
	TokNumExp
	TokNumFracNum
	TokNumFracSign
	TokNumFracDenom
	TokNumPct
	TokSpace  // insert a space the width of the specified character
	TokRepeat // repeat the specified character to fill the width of the column
	TokCondition

	TokAMPM
	TokMonth
	TokDay
	TokYear
	TokHour
	TokMinute
	TokSecond
	TokSecFraction
	TokTotalHours
	TokTotalMinutes
	TokTotalSeconds
	TokLiteral
)

type FmtToken struct {
	Type FmtTokenType
	Size int
	Data string
}

type FormatType int

const (
	NoType FormatType = iota
	TextFormat
	TimeFormat
	NumberFormat
	BoolFormat
	ErrorFormat
)

type FormatSubType int

const (
	NoSubType = iota
	DateTime
	Date
	Time
	Duration
)

type Section struct {
	Type    FormatType
	SubType FormatSubType
	Tokens  []FmtToken
}

/*
func (ft *Section) IsDuration() bool {
	if ft.Type == TimeFormat {
		for _, t := range ft.Tokens {
			if t.Type == TokTotalHours || t.Type == TokTotalMinutes || t.Type == TokTotalSeconds {
				return true
			}
		}
	}
	return false
}
*/

// read a sequence of bytes of teh same character, returning the count and the remaining data.
func readRepeat(input []byte) (rem []byte, ch byte, count int) {
	ch, input = input[0], input[1:]
	count = 1
	for len(input) > 0 && input[0] == ch {
		count++
		input = input[1:]
	}
	return input, ch, count
}

// read characters matching those listed in chars from input; returns remaining data after first non-match.
func readChars(input, chars []byte) (out, rem []byte) {
	var i int
SCAN:
	for i = 0; i < len(input); i++ {
		for _, ch := range chars {
			if input[i] == ch {
				continue SCAN
			}
		}
		break
	}
	return input[0:i], input[i:]
}

func LastTokenIdxByType(tokens []FmtToken, findType FmtTokenType) (idx int) {
	for i := len(tokens) - 1; i >= 0; i-- {
		if tokens[i].Type == findType {
			return i
		}
	}
	return -1
}

func prevTokenType(tokens []FmtToken, i int) FmtTokenType {
	for i--; i >= 0; i-- {
		switch tokens[i].Type {
		case TokLiteral, TokSpace, TokRepeat:
		default:
			return tokens[i].Type
		}
	}
	return TokInvalid
}

func nextTokenType(tokens []FmtToken, i int) FmtTokenType {
	for i++; i < len(tokens); i++ {
		switch tokens[i].Type {
		case TokLiteral, TokSpace, TokRepeat:
		default:
			return tokens[i].Type
		}
	}
	return TokInvalid
}

// read a string up to a terminating chracter, skipping escaped instances
// of that character (and stripping the escape character itself)
func readToChar(input []byte, ch byte) (txt, rem []byte) {
	for len(input) > 0 {
		i := bytes.IndexByte(input, ch)
		if i < 0 {
			// not found
			return nil, input
		}
		if i == 0 || input[i-1] != '\\' {
			txt = append(txt, input[:i]...)
			return txt, input[i+1:]
		}
		txt = append(txt, input[:i-1]...) // skip backslash
		txt = append(txt, input[i])       // Add escaped character
		input = input[i+1:]               // skip backslash
	}
	return nil, input
}

func skipEscape(input []byte) (txt, rem []byte) {
	if len(input) > 1 {
		return input[1:2], input[2:] // skip backslsah
	}
	return input[0:0], input[0:0]
}

func flush(tokens []FmtToken, other []byte) ([]FmtToken, []byte) {
	if len(other) > 0 {
		tokens = append(tokens, FmtToken{Type: TokLiteral, Data: string(other)})
		other = other[0:0]
	}
	return tokens, other
}

//func tokenizeTime(current []FmtToken, format []byte) (tokens []FmtToken, rem []byte) {
func tokenizeTime(tokens []FmtToken, format []byte) (section Section, rem []byte) {
	var (
		other                      []byte
		count                      int
		ft                         FmtTokenType
		hasDate, hasTime, hasTotal bool
	)

	rem = format
TLOOP:
	for len(rem) > 0 {
		switch ch := rem[0]; ch {
		case 'y', 'm', 'd', 'h', 's', '0':
			rem, ch, count = readRepeat(rem)
			switch ch {
			case 'y':
				ft = TokYear
				hasDate = true
			case 'm':
				ft = TokMonth
				// may be changed to minute later, depending on context
			case 'd':
				ft = TokDay
				hasDate = true
			case 'h':
				ft = TokHour
				hasTime = true
			case 's':
				ft = TokSecond
				hasTime = true
			case '0':
				ft = TokSecFraction
				hasTime = true
			default:
				panic("Unhandled character token")
			}

			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: ft, Size: count})

		case '[':
			// read total hour/minute/second token [h], [m], [s]
			if len(rem) < 3 {
				other = append(other, '[')
				rem = rem[1:]
				break
			}
			org := rem
			rem, ch, count = readRepeat(rem[1:])
			if len(rem) < 1 || rem[0] != ']' {
				other = append(other, '[')
				rem = org[1:]
				break

			}
			switch ch {
			case 'h':
				ft = TokTotalHours
				hasTotal = true
			case 'm':
				ft = TokTotalMinutes
				hasTotal = true
			case 's':
				ft = TokTotalSeconds
				hasTotal = true
			default:
				// ignore; can't be a color or conditional as excel always places those
				// at the start of the section; shouldn't occur in a legal format
				other = append(other, '[')
				rem = org[1:]
				continue
			}
			rem = rem[1:] // skip ']'
			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: ft, Size: count})

		case 'A', 'a':
			if len(rem) >= 3 && (string(rem[:3]) == "A/P" || string(rem[:3]) == "a/p") {
				tokens, other = flush(tokens, other)
				tokens = append(tokens, FmtToken{Type: TokAMPM, Size: 1, Data: string(rem[0:1])})
				rem = rem[3:]
			} else if len(rem) >= 5 && (string(rem[:5]) == "AM/PM" || string(rem[:5]) == "am/pm") {
				tokens, other = flush(tokens, other)
				tokens = append(tokens, FmtToken{Type: TokAMPM, Size: 2, Data: string(rem[0:1])})
				rem = rem[5:]
			} else {
				other = append(other, rem[0])
				rem = rem[1:]
			}

		case ';':
			// named semi-colon ends the section
			// TODO
			tokens, other = flush(tokens, other)
			rem = rem[1:] // drop semi-colon
			break TLOOP
			//return tokens, rem[1:] // drop the semi-colon

		default:
			tokens, rem, other = tokenizeCommon(tokens, rem, other)
		}
	}
	tokens, other = flush(tokens, other)

	// 'm' is either month or minute depending on context
	for i, token := range tokens {
		if token.Type == TokMonth {
			if pt, nt := prevTokenType(tokens, i), nextTokenType(tokens, i); pt == TokHour || pt == TokTotalHours || nt == TokSecond || nt == TokTotalSeconds {
				tokens[i].Type = TokMinute
				hasTime = true
			} else {
				hasDate = true
			}
		}
	}

	section = Section{
		Type:   TimeFormat,
		Tokens: tokens,
	}
	switch {
	case hasTotal:
		section.SubType = Duration
	case hasDate && hasTime:
		section.SubType = DateTime
	case hasDate:
		section.SubType = Date
	case hasTime:
		section.SubType = Time
	default:
		section.SubType = NoSubType
	}

	return section, rem
}

/*
1234.59  -> ####.# -> 1234.6
8.9      -> #.000 -> 8.900
.631     -> 0.# -> 0.6
12       -> #.0# -> 12.0
1234.568 -> #.0# -> 123.57
2.8      -> ???.??? -> 2.8 (align decimal)
5.25     -> # ???/??? -> 5 1/4 (aligned fractions)
12000    -> #,### -> 12,000
12000    -> #, -> 12  (divide by 1000)
12200000 -> 0.0,, -> 12.2 (divide by 1000 x 10000)
2.8      -> 0.0% -> 280%
123000   -> #.#E+00 -> 1.2E+05
123000   -> #.#E-00 -> 1.2E05
123000   -> #.#E00 -> 1.2E+05

0 -> displays zeroes for insignificant sigits
# -> displays only signficant digits
? -> displays spaces for padding for insignificant digits

number (format, #.0# etc inc percentage)
exponent
fraction

also: "_" includes a space the width of one character
      "*" causes the next cahracter to be repeated

	# # /##
*/
//func tokenizeNumeric(current []FmtToken, format []byte) (tokens []FmtToken, rem []byte) {
func tokenizeNumeric(tokens []FmtToken, format []byte) (section Section, rem []byte) {
	var other []byte
	var seenInt bool
	var inFrac bool
	var inDec bool
	_, _ = inDec, seenInt

	rem = format
NLOOP:
	for len(rem) > 0 {
		switch ch := rem[0]; ch {
		case '0', '?', '#':
			// start of number pattern for integer, decimal, numerator or denominator
			var num []byte
			var ttype FmtTokenType
			switch {
			case inFrac:
				// denominator
				num, rem = readChars(rem, []byte("0?#"))
				ttype = TokNumFracDenom
				inFrac = false
			case inDec:
				num, rem = readChars(rem, []byte("0?#,"))
				ttype = TokNumDec
				inDec = false
			default:
				num, rem = readChars(rem, []byte("0?#,"))
				ttype = TokNumInt
			}

			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: ttype, Data: string(num)})

		case '.':
			tokens = append(tokens, FmtToken{Type: TokNumDecSign})
			inDec = true
			rem = rem[1:]

		case '/':
			// find the previous integer token and turn it into the numerator
			if idx := LastTokenIdxByType(tokens, TokNumInt); idx > -1 {
				tokens, other = flush(tokens, other)
				tokens[idx].Type = TokNumFracNum
				tokens = append(tokens, FmtToken{Type: TokNumFracSign})
				inFrac = true
			}
			rem = rem[1:]

		case '1', '2', '3', '4', '5', '6', '7', '8', '9':
			// possibly a fractional denominator
			var num []byte
			if inFrac {
				num, rem = readChars(rem, []byte("0123456789"))
				tokens, other = flush(tokens, other)
				tokens = append(tokens, FmtToken{Type: TokNumFracDenom, Data: string(num)})
				inFrac = false
			} else {
				// treat as text
				num, rem = readChars(rem, []byte("0123456789.,"))
				other = append(other, num...)
			}

		case '%':
			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: TokNumPct})
			rem = rem[1:]

		case 'E', 'e':
			// read exponent definition
			var s []byte
			exp := rem[0:1]
			tokens, other = flush(tokens, other)
			s, rem = readChars(rem[1:], []byte("?#0+"))
			tokens = append(tokens, FmtToken{Type: TokNumExp, Data: string(append(exp, s...))})

		case ';':
			// named semi-colon ends the section
			tokens, other = flush(tokens, other)
			rem = rem[1:] // drop semi-colon
			break NLOOP
			//return tokens, rem[1:] // drop the semi-colon

		default:
			// handle quoted strings, etc
			tokens, rem, other = tokenizeCommon(tokens, rem, other)
		}
	}

	tokens, other = flush(tokens, other)
	section = Section{Type: NumberFormat, Tokens: tokens}
	return section, rem
}

// handle tokens that are common to both times and number sections
func tokenizeCommon(tokens []FmtToken, rem, other []byte) ([]FmtToken, []byte, []byte) {
	switch ch := rem[0]; ch {
	case '\\':
		// escape character
		var txt []byte
		txt, rem = skipEscape(rem)
		other = append(other, txt...)

	case '"':
		// quoted string
		var quoted []byte
		quoted, rem = readToChar(rem[1:], '"')
		other = append(other, quoted...)

	case '_', '*':
		// Underscore + character = a blank space the width of the character
		// Star + character = repeat character to fill up the width of the cell
		if len(rem) > 1 {
			tokType := TokSpace
			if rem[0] == '*' {
				tokType = TokRepeat
			}
			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: tokType, Data: string(rem[1:2])})
			rem = rem[2:]
		} else {
			other = append(other, '_')
			rem = rem[1:]
		}

	default:
		// treat anything else as literal text
		other = append(other, ch)
		rem = rem[1:]
	}
	return tokens, rem, other
}

func isColor(col []byte) bool {
	c := strings.ToLower(string(col))
	switch c {
	// per https://support.office.com/en-nz/article/Create-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4?ui=en-US&rs=en-NZ&ad=NZ
	case "black", "green", "white", "blue", "magenta", "yellow", "cyan", "red":
		return true
	}
	return strings.HasPrefix(c, "color")
}

type CellFormat struct {
	Sections      []Section
	IsConditional bool // true if the choice of token section is based on a condition
}

type FormattedValue struct {
	GoValue        interface{} // string, float64, bool, time.Time or time.Duration
	FormattedValue string
	//FormatType     FormatType
	//FormatSubType  FormatSubType
	Section Section // Format section that was used to generated FormattedValue
}

// ParseFormat parses the format string for a cell into a CellFormat that can be applied
// to cell values.
func ParseFormat(format string) (t CellFormat) {
	var other []byte
	var sets []Section
	var tokens []FmtToken
	var section Section
	var conditional bool
	rem := []byte(format)

	// Decide if we're tokenizing a date/duration or a number
	for len(rem) > 0 {
		switch ch := rem[0]; ch {
		case 'y', 'm', 'd', 'h', 's': // don't include "0" here; not valid for time without a prefix
			// time
			tokens, other = flush(tokens, other)
			section, rem = tokenizeTime(tokens, rem)
			sets = append(sets, section)
			//sets = append(sets, Section{Type: TimeFormat, Tokens: tokens})
			tokens = nil

		case '#', '?', '0', '.':
			// numeric
			tokens, other = flush(tokens, other)
			section, rem = tokenizeNumeric(tokens, rem)
			//sets = append(sets, Section{NumberFormat, tokens})
			sets = append(sets, section)
			tokens = nil

		case 'g', 'G':
			if len(rem) >= len("general") && strings.ToLower(string(rem[:len("general")])) == "general" {
				tokens, other = flush(tokens, other)
				tokens = append(tokens, FmtToken{Type: TokGeneral})
				rem = rem[len("general"):]
			} else {
				other = append(other, ch)
				rem = rem[1:]
			}

		case '[':
			// could be time related, a colour, or a condition
			entry, rem2 := readToChar(rem[1:], ']')
			switch {
			case len(entry) == 0:
				// no closing bracket?
				// swallow

			case isColor(entry):
				tokens, other = flush(tokens, other)
				tokens = append(tokens, FmtToken{Type: TokColor, Data: string(entry)})

			default:
				switch entry[0] {
				case 'h', 'm', 's':
					// time
					// TODO: validate that all character in entry are the same (eg. hhh)
					tokens, other = flush(tokens, other)
					section, rem2 = tokenizeTime(tokens, rem)
					sets = append(sets, section)
					//sets = append(sets, Section{TimeFormat, tokens})
					tokens = nil

				case '<', '>', '=':
					// condition
					tokens, other = flush(tokens, other)
					tokens = append(tokens, FmtToken{Type: TokCondition, Data: string(entry)})
					conditional = true

				default:
					// it's technically invalid afaik; swallow it
				}
			}
			rem = rem2

		case ';':
			// end section
			tokens, other = flush(tokens, other)
			sets = append(sets, Section{Type: TextFormat, Tokens: tokens})
			tokens = nil
			rem = rem[1:]

		case '@':
			// insert literal cell text
			tokens, other = flush(tokens, other)
			tokens = append(tokens, FmtToken{Type: TokCellText})
			rem = rem[1:]

		default:
			// handle quoted strings, etc
			tokens, rem, other = tokenizeCommon(tokens, rem, other)
		}
	}

	tokens, other = flush(tokens, other)
	if len(tokens) > 0 {
		sets = append(sets, Section{Type: TextFormat, Tokens: tokens})
	}

	t.Sections = sets
	t.IsConditional = conditional

	return t
}

// FormatValue applies the parsed format to a specified value.
// This is usually accessed via a Cell.
/*
func (ct CellFormat) FormatValue(sv string, cellType CellType, date1904 bool) (s string, err error) {
	var fv float64
	var isNumber bool

	switch cellType {

	case CellTypeBool:
		// Not sure bools can be formatted?
		if sv == "1" {
			return "TRUE", nil
		}
		return "FALSE", nil

	case CellTypeError:
		return sv, nil

	case CellTypeNumeric, CellTypeFormula:
		isNumber = true

	case CellTypeString, CellTypeInline:
		// string
		isNumber = false

	default:
		panic("Unhandled cell type")

	}

	if isNumber {
		fv, err = strconv.ParseFloat(sv, 64)
		if err != nil {
			return "", err
		}
	}

	if ct.IsConditional {
		return "", UnsupportedCondition
	}

	// note; the order of these cases  is important!
	switch scount := len(ct.Sections); {
	case scount < 4 && !isNumber:
		// text, but no text section
		return sv, nil

	case scount == 0:
		// TODO should format this if its a number
		return sv, nil

	case scount == 1 || (scount == 2 && fv == 0) || fv > 0:
		// positive entry only
		// always used if its the only section or the value is positive,
		// else used if there's two sections but the value is zero
		// value is implicitly not text by this case
		formatted, err := formatValue(ct.Sections[0], sv, math.Abs(fv), date1904)
		if err != nil {
			return "", err
		}
		if fv < 0 {
			return "-" + formatted, nil
		}
		return formatted, nil

	case scount >= 2 && fv < 0:
		// positive & negative values
		// fv is implicitly negative here; if it were positive it would of been caught above
		return formatValue(ct.Sections[1], sv, math.Abs(fv), date1904)

	case scount >= 3 && fv == 0 && isNumber:
		// positive ;negative; zero
		// fv is implicitly zero
		return formatValue(ct.Sections[2], sv, fv, date1904)

	case scount > 3 && !isNumber:
		return formatValue(ct.Sections[3], sv, math.Abs(fv), date1904)

	default:
		panic("Unhandled condition") // Bug if this occurs!
	}
}
*/

func (ct CellFormat) FormatValue(sv string, cellType CellType, date1904 bool) (v FormattedValue, err error) {
	var fv float64
	var isNumber bool

	switch cellType {

	case CellTypeBool:
		// Not sure bools can be formatted?
		v.Section = Section{Type: BoolFormat}
		if sv == "1" {
			v.GoValue = true
			v.FormattedValue = "TRUE"
		} else {
			v.GoValue = false
			v.FormattedValue = "FALSE"
		}
		return v, nil

	case CellTypeError:
		v.Section = Section{Type: ErrorFormat}
		v.GoValue = sv
		v.FormattedValue = sv
		return v, nil

	case CellTypeNumeric, CellTypeFormula:
		isNumber = true

	case CellTypeString, CellTypeInline:
		// string
		isNumber = false

	default:
		panic("Unhandled cell type")

	}

	if isNumber {
		fv, err = strconv.ParseFloat(sv, 64)
		if err != nil {
			return v, err
		}
	}

	if ct.IsConditional {
		return v, UnsupportedCondition
	}

	// note; the order of these cases  is important!
	switch scount := len(ct.Sections); {
	case scount < 4 && !isNumber:
		// text, but no text section
		v.Section = Section{Type: TextFormat}
		v.GoValue = sv
		v.FormattedValue = sv
		return v, nil

	case scount == 0:
		// TODO should format this if its a number
		// implicitly a number
		v.Section = Section{Type: NumberFormat}
		v.GoValue = fv
		v.FormattedValue = sv
		return v, nil

	case scount == 1 || (scount == 2 && fv == 0) || fv > 0:
		// positive entry only
		// always used if its the only section or the value is positive,
		// else used if there's two sections but the value is zero
		// value is implicitly not text by this case
		v.GoValue = fv
		err := formatValue(&v, ct.Sections[0], sv, math.Abs(fv), date1904)
		if err != nil {
			return v, err
		}
		if fv < 0 {
			v.FormattedValue = "-" + v.FormattedValue
			return v, nil
			//return "-" + formatted, nil
		}
		return v, nil

	case scount >= 2 && fv < 0:
		// positive & negative values
		// fv is implicitly negative here; if it were positive it would of been caught above
		v.GoValue = fv
		err := formatValue(&v, ct.Sections[1], sv, math.Abs(fv), date1904)
		return v, err

	case scount >= 3 && fv == 0 && isNumber:
		// positive ;negative; zero
		// fv is implicitly zero
		v.GoValue = fv
		err := formatValue(&v, ct.Sections[2], sv, fv, date1904)
		return v, err

	case scount > 3 && !isNumber:
		v.GoValue = sv
		err := formatValue(&v, ct.Sections[3], sv, math.Abs(fv), date1904)
		return v, err

	default:
		panic("Unhandled condition") // Bug if this occurs!
	}
}

/*
func (ct CellFormat) FormatType(sv string, cellType CellType, date1904 bool) (FormatType, FormatSubType, error) {
	var fv float64
	var err error

	switch cellType {
	case CellTypeBool:
		return FormatTypeBool NoSubType, nil
	}
	if isNumber {
		fv, err = strconv.ParseFloat(sv, 64)
		if err != nil {
			return NoType, NoSubType, err
		}
	}

	if ct.IsConditional {
		return NoType, NoSubType, UnsupportedCondition
	}

	// note; the order of these cases  is important!
	switch scount := len(ct.Sections); {
	case scount < 4 && !isNumber:
		// text, but no text section
		return TextFormat, NoSubType, nil

	case scount == 0:
		if isNumber {
			return NumberFormat, NoSubType, nil
		}
		return TextFormat, NoSubType, nil

	case scount == 1 || (scount == 2 && fv == 0) || fv > 0:
		return ct.Sections[0].Type, ct.Sections[0].SubType, nil

	case scount >= 2 && fv < 0:
		// positive & negative values
		// fv is implicitly negative here; if it were positive it would of been caught above
		return ct.Sections[1].Type, ct.Sections[1].SubType, nil

	case scount >= 3 && fv == 0 && isNumber:
		// positive ;negative; zero
		// fv is implicitly zero
		return ct.Sections[2].Type, ct.Sections[2].SubType, nil

	case scount > 3 && !isNumber:
		return ct.Sections[3].Type, ct.Sections[3].SubType, nil

	default:
		panic("Unhandled condition") // Bug if this occurs!
	}
}
*/

//func formatValue(v *FormattedValue, s Section, sv string, fv float64, date1904 bool) (string, error) {
func formatValue(v *FormattedValue, s Section, sv string, fv float64, date1904 bool) error {
	v.Section = s
	switch s.Type {
	case TimeFormat:
		fstr, t, d := formatTime(s.Tokens, fv, date1904)
		v.FormattedValue = fstr
		// override numeric value
		if s.SubType == Duration {
			fmt.Println("SET DURATION", d)
			v.GoValue = d
		} else {
			fmt.Println("SET TIME", t)
			v.GoValue = t
		}
		return nil

	case NumberFormat:
		fstr := formatNumber(s.Tokens, fv)
		v.FormattedValue = fstr
		return nil

	case TextFormat:
		fstr := formatText(s.Tokens, sv, fv)
		v.FormattedValue = fstr
		return nil

	default:
		panic(fmt.Sprintf("Unknown token type %d", s.Type))
	}
}

func formatText(tokens []FmtToken, sv string, fv float64) string {
	var output []byte

	for _, token := range tokens {
		switch token.Type {
		case TokCellText:
			output = append(output, sv...)

		case TokGeneral:
			// TODO: this should format the number, rather than insert literal text
			output = append(output, strconv.FormatFloat(fv, 'f', -1, 64)...)

		case TokLiteral:
			output = append(output, token.Data...)
		}
	}

	return string(output)
}

func formatTime(tokens []FmtToken, v float64, date1904 bool) (string, time.Time, time.Duration) {
	var (
		f      string
		output []byte
		res    = time.Second // round to nearest second by default
	)

	d := DurationFromExcelTime(v)
	t := TimeFromExcelTime(v, date1904)
	h1fmt, h2fmt := "15", "15" // 24 hour time

	for _, token := range tokens {
		switch token.Type {
		case TokAMPM:
			h1fmt, h2fmt = "3", "03"
		case TokSecFraction:
			res = time.Second / time.Duration(math.Pow10(token.Size))
		}
	}

	t = t.Round(res) // round to second or finer resolution

	for _, token := range tokens {
		switch token.Type {
		case TokYear:
			if token.Size > 2 {
				f = t.Format("2006")
			} else {
				f = t.Format("06")
			}

		case TokMonth:
			switch token.Size {
			case 1:
				f = t.Format("1")
			case 2:
				f = t.Format("01")
			case 3:
				f = t.Format("Jan")
			case 5:
				f = t.Format("Jan")[0:1]
			default:
				f = t.Format("January")
			}

		case TokDay:
			switch token.Size {
			case 1:
				f = t.Format("2")
			case 2:
				f = t.Format("02")
			case 3:
				f = t.Format("Mon")
			default:
				f = t.Format("Monday")
			}

		case TokHour:
			switch token.Size {
			case 1:
				f = t.Format(h1fmt)
			default:
				f = t.Format(h2fmt)
			}

		case TokAMPM:
			switch token.Size {
			case 1:
				f = t.Format("PM")[0:1]
			default:
				f = t.Format("PM")
			}
			if token.Data == "a" {
				f = strings.ToLower(f)
			}

		case TokMinute:
			switch token.Size {
			case 1:
				f = t.Format("4")
			default:
				f = t.Format("04")
			}

		case TokSecond:
			switch token.Size {
			case 1:
				f = t.Format("5")
			default:
				f = t.Format("05")
			}

		case TokSecFraction:
			f = t.Format("." + strings.Repeat("0", token.Size))[1:]

		case TokTotalHours:
			f = string(padint(token.Size, int64(d/time.Hour)))

		case TokTotalMinutes:
			f = string(padint(token.Size, int64(d/time.Minute)))

		case TokTotalSeconds:
			f = string(padint(token.Size, int64(d/time.Second)))

		case TokLiteral:
			f = token.Data
		}

		output = append(output, f...)
	}

	return string(output), t, d
}

// formatNumber takes a postiive value and formats it per the token set
// (negative numbers are converted to positive during the FormatValue caller)
func formatNumber(tokens []FmtToken, v float64) string {
	var (
		output             []string
		intFmt             []byte
		decFmt             []byte
		expPrec            int = -1
		decPrec            int
		commaCount         int
		hasComma           bool // use thousands separator
		hasInt, hasExp     bool
		fracNum, fracDenom int64
		fracDenomFmt       string
	)

	// scan tokens
	for _, token := range tokens {
		switch token.Type {
		case TokNumPct:
			// percent operator causes v to be multiplied by 100
			v *= 100

		case TokNumInt:
			hasInt = true
			intFmt = make([]byte, 0, len(token.Data))

			// strip/count trailing commas
			data := stripTrailingComma(&v, token.Data)
			fmt.Println("data", data)
			for _, ch := range []byte(data) { // XXX utf8 issues?
				switch ch {
				case ',':
					// non-trailing comma
					hasComma = true
				default:
					intFmt = append(intFmt, ch)
				}
			}
			fmt.Printf("inFmt now %q\n", string(intFmt))

		case TokNumDec:
			// data only contains 0?#,
			decFmt = make([]byte, 0, len(token.Data))
			var c int
			for _, ch := range stripTrailingComma(&v, token.Data) {
				switch ch {
				case '#', '0', '?', '.':
					if ch != '.' {
						decPrec++
						c++
					}
					decFmt = append(decFmt, ch)
				default:
					c = 0
				}
			}
			commaCount += c

		case TokNumFracSign:
			decPrec = 1 // ensure integer isn't rounded off
			break

		case TokNumFracDenom:
			fracDenomFmt = token.Data

		case TokNumExp:
			// E+00, E00, e-00	kk
			fmt.Println("EXP", token.Size)
			hasExp = true
			expPrec = token.Size
		}
	}

	if fracDenomFmt != "" {
		// format the numerator and denominator for future rendering
		// determime if an exact denominator is required
		var err error
		f := v
		if hasInt {
			// extract floating point portion if integer already displayed
			_, f = math.Modf(v)
		}
		if ch := fracDenomFmt[0]; ch >= '1' && ch <= '9' {
			// find closest fraction for exact denmoniator
			fracDenom, err = strconv.ParseInt(fracDenomFmt, 10, 64)
			if err == nil {
				fracNum = int64(math.Floor((float64(fracDenom) * f) + 0.5))
			}
		} else {

			fmt.Printf("FRAP f=%f, md=%d\n", math.Abs(f), int64(math.Pow(10, float64(len(fracDenomFmt)))-1))
			fracNum, fracDenom = frap(math.Abs(f), int64(math.Pow(10, float64(len(fracDenomFmt)))-1))
		}
	}

	// format the number
	intval, decval, expval := splitNum(v, expPrec, decPrec)
	fmt.Println("splitout", intval, decval, expval)

	pretty.Println("Tokens", tokens)

	for _, token := range tokens {
		switch token.Type {
		case TokNumInt:
			// integer portion of number
			if intval == "0" && !hasExp {
				intval = ""
			}
			fmt.Println("int", string(intFmt), intval, len(intFmt), len(intval))
			fmt.Printf("intfmt=%q intval=%q\n", intFmt, intval)
			intval, sigonly := formatInteger(intFmt, intval)
			if !(sigonly && intval == "0" && !hasExp) {
				if hasComma {
					output = append(output, fmtThou(intval))
				} else {
					output = append(output, intval)
				}
			}

		case TokNumDecSign:
			output = append(output, ".")

		case TokNumDec:
			// decimal portion of number
			output = append(output, formatDecimal(decFmt, decval))

		case TokNumExp:
			intval := strings.TrimLeft(expval[2:], "0") // strip "E+" and leading zeroes
			if intval == "" {
				intval = "0"
			}
			v, _ := formatInteger([]byte(token.Data[2:]), intval)
			output = append(output, "E+")
			output = append(output, v)

		case TokNumFracSign:
			output = append(output, "/")

		case TokNumFracNum:
			// fractional numerator
			_, numval := fmtSig(strconv.FormatInt(fracNum, 10), token.Data)
			output = append(output, numval)

		case TokNumFracDenom:
			// fractional denominator
			_, denomval := fmtSig(strconv.FormatInt(fracDenom, 10), token.Data)
			output = append(output, denomval)

		case TokNumPct:
			output = append(output, "%")

		case TokLiteral:
			output = append(output, token.Data)
		}
	}

	return strings.Join(output, "")
}

func formatInteger(intfmt []byte, intval string) (v string, sigonly bool) {
	var prefix []byte
	sigonly = true
	fmt.Printf("infmt=%q intval=%q\n", string(intfmt), intval)
	for i := 0; i < len(intfmt)-len(intval); i++ {
		fmt.Printf("int ch i=%d ch=%c\n", i, intfmt[i])
		switch ch := intfmt[i]; ch {
		case '0':
			prefix = append(prefix, '0')
			sigonly = false
		case '?':
			prefix = append(prefix, ' ')
			sigonly = false
		}
	}
	return string(append(prefix, intval...)), sigonly
}

func formatDecimal(decfmt []byte, decval string) string {
	fmt.Printf("decfmt=%q decval=%q\n", decfmt, decval)
	dvl := len(decval)
	for i := 0; i < len(decfmt)-dvl; i++ {
		fmt.Println("FMT CH", decfmt[i+dvl])
		switch ch := decfmt[i+dvl]; ch {
		case '0':
			fmt.Println("add zero")
			decval += "0"
		case '?':
			decval += " "
		}
	}
	fmt.Println("formatDec result", decval)
	return decval
}

func fmtThou(intval string) string {
	outpos := 2 * len(intval)
	out := make([]byte, 2*len(intval))
	for i := len(intval); i > 0; i -= 3 {
		p := i - 3
		if p > 0 {
			copy(out[outpos-3:], intval[p:p+3])
			out[outpos-4] = ','
			outpos -= 4

		} else {
			copy(out[outpos-i:], intval[0:i])
			outpos -= i
		}
	}

	return string(out[outpos:])
}

func fmtSig(intval string, ifmt string) (sigonly bool, out string) {
	var prefix []byte
	sigonly = true
	fmt.Println("int", string(ifmt), intval, len(ifmt), len(intval))
	for i := 0; i < len(ifmt)-len(intval); i++ {
		fmt.Println("ch", i, ifmt[i])
		switch ch := ifmt[i]; ch {
		case '0':
			prefix = append(prefix, '0')
			sigonly = false
		case '?':
			prefix = append(prefix, ' ')
			sigonly = false
		}
	}
	return sigonly, string(append(prefix, intval...))
}

func splitNum(v float64, expPrec, decPrec int) (intval, decval, expval string) {
	var s string
	if expPrec >= 0 {
		// one day we'll do something with expPrec other than treating it as a bool
		s = strconv.FormatFloat(v, 'E', decPrec, 64)
		idx := strings.IndexByte(s, 'E')
		expval = s[idx:]
		s = s[:idx]
	} else {
		s = strconv.FormatFloat(v, 'f', decPrec, 64)
	}
	if decPrec > 0 {
		idx := strings.IndexByte(s, '.')
		intval = s[:idx]
		// trim trailing zeroes
		for decval = s[idx+1:]; len(decval) > 0 && decval[len(decval)-1] == '0'; decval = decval[:len(decval)-1] {
		}
	} else {
		intval = s
	}
	return intval, decval, expval
}

// trip trailing commas from format and divide v by 1000 for each one found
func stripTrailingComma(v *float64, fmt string) (stripped []byte) {
	var end int
	for end = len(fmt) - 1; end >= 0 && fmt[end] == ','; end-- {
		*v /= 1000
	}
	return []byte(fmt[0 : end+1])
}

func padint(size int, val int64) (v []byte) {
	v = strconv.AppendInt(v, val, 10)
	if lv := len(v); lv < size {
		r := append([]byte{}, "0000000000000000000"[:size-lv]...)
		v = append(r, v...)
	}
	return v
}

// frap finds rational approximation to given real number
//
// pretty literal conversion of https://www.ics.uci.edu/~eppstein/numth/frap.c
// by David Eppstein
func frap(n float64, maxDenom int64) (num, denom int64) {
	var ai int64
	m := [2][2]int64{{1, 0}, {0, 1}}
	x := n

	for ai = int64(n); m[1][0]*ai+m[1][1] < maxDenom; ai = int64(x) {
		t := m[0][0]*ai + m[0][1]
		m[0][1] = m[0][0]
		m[0][0] = t
		t = m[1][0]*ai + m[1][1]
		m[1][1] = m[1][0]
		m[1][0] = t
		if x == float64(ai) {
			break
		}
		x = 1 / (x - float64(ai))
		if x > float64(math.MaxFloat64) {
			break
		}
	}

	num, denom = m[0][0], m[1][0]
	err1 := n - (float64(m[0][0]) / float64(m[1][1]))
	if err1 == 0 {
		return num, denom
	}

	ai = (maxDenom - m[1][1]) / m[1][0]
	m[0][0] = m[0][0]*ai + m[0][1]
	m[1][0] = m[1][0]*ai + m[1][1]

	err2 := n - (float64(m[0][0]) / float64(m[1][0]))

	if err1 < err2 {
		return num, denom
	}
	return m[0][0], m[1][0]
}
