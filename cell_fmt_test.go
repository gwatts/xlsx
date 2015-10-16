package xlsx

import (
	"fmt"
	"time"

	. "gopkg.in/check.v1"
)

type CellFmtSuite struct{}

var _ = Suite(&CellFmtSuite{})

var readCharTests = []struct {
	input    string
	expected string
	rem      string
}{
	{"", "", ""},
	{"a", "a", ""},
	{"agh", "a", "gh"},
	{"abgh", "ab", "gh"},
	{"gh", "", "gh"},
}

func (s *CellFmtSuite) TestReadChars(c *C) {
	for _, test := range readCharTests {
		out, rem := readChars([]byte(test.input), []byte("ab"))
		c.Assert(string(out), Equals, test.expected, Commentf("input=%q", test.input))
		c.Assert(string(rem), Equals, test.rem, Commentf("input=%q", test.input))
	}
}

var escTests = []struct {
	input string
	esc   string
	rem   string
}{
	{"\\foo", "f", "oo"},
	{"\\f", "f", ""},
	{"\\", "", ""},
}

func (s *CellFmtSuite) TestSkipEscape(c *C) {
	for _, test := range escTests {
		txt, rem := skipEscape([]byte(test.input))
		c.Assert(string(txt), Equals, test.esc, Commentf("txt mismatch input=%q", test.input))
		c.Assert(string(rem), Equals, test.rem, Commentf("rem mismatch input=%q", test.input))
	}
}

var readToCharTests = []struct {
	input string
	txt   string
	rem   string
}{
	{``, ``, ``},
	{`no term`, ``, `no term`},
	{`quoted"`, `quoted`, ``},
	{`quoted" string`, `quoted`, ` string`},
	{`quoted esc\"aped" string`, `quoted esc"aped`, ` string`},
}

func (s *CellFmtSuite) TestReadToChar(c *C) {
	for _, test := range readToCharTests {
		txt, rem := readToChar([]byte(test.input), '"')
		c.Assert(string(txt), Equals, test.txt, Commentf("txt mismatch input=%q", test.input))
		c.Assert(string(rem), Equals, test.rem, Commentf("rem mismatch input=%q", test.input))
	}
}

func (s *CellFmtSuite) TestReadRepeat(c *C) {
	input := []byte("aaabbc")
	expected := []string{"aaa", "bb", "c"}
	var ch byte
	var count int
	for _, ex := range expected {
		input, ch, count = readRepeat(input)
		c.Assert(ch, Equals, ex[0])
		c.Assert(count, Equals, len(ex))
	}
	c.Assert(len(input), Equals, 0)
}

var timeTokenizeTests = []struct {
	input    string
	rem      string
	expected Section
}{
	{"h", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{{TokHour, 1, ""}}}},
	{"h:mm am/pm", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{
		{TokHour, 1, ""},
		{TokLiteral, 0, ":"},
		{TokMinute, 2, ""},
		{TokLiteral, 0, " "},
		{TokAMPM, 2, "a"}}}},
	{"h:mm a/p", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{
		{TokHour, 1, ""},
		{TokLiteral, 0, ":"},
		{TokMinute, 2, ""},
		{TokLiteral, 0, " "},
		{TokAMPM, 1, "a"}}}},
	{"mm:ss.00", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{
		{TokMinute, 2, ""},
		{TokLiteral, 0, ":"},
		{TokSecond, 2, ""},
		{TokLiteral, 0, "."},
		{TokSecFraction, 2, ""}}}},
	{"yy-mm-dd", "", Section{Type: TimeFormat, SubType: Date, Tokens: []FmtToken{
		{TokYear, 2, ""},
		{TokLiteral, 0, "-"},
		{TokMonth, 2, ""},
		{TokLiteral, 0, "-"},
		{TokDay, 2, ""}}}},
	{"hh[xz]mm", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{
		{TokHour, 2, ""},
		{TokLiteral, 0, "[xz]"},
		{TokMinute, 2, ""}}}},
	{"[hh]:[mm]:[ss]", "", Section{Type: TimeFormat, SubType: Duration, Tokens: []FmtToken{
		{TokTotalHours, 2, ""},
		{TokLiteral, 0, ":"},
		{TokTotalMinutes, 2, ""},
		{TokLiteral, 0, ":"},
		{TokTotalSeconds, 2, ""}}}},
	{"foobar", "", Section{Type: TimeFormat, SubType: NoSubType, Tokens: []FmtToken{
		{TokLiteral, 0, "foobar"}}}},
	{"foo;bar", "bar", Section{Type: TimeFormat, SubType: NoSubType, Tokens: []FmtToken{
		{TokLiteral, 0, "foo"}}}},
	{"\"skip;semi\" foo;bar", "bar", Section{Type: TimeFormat, SubType: NoSubType, Tokens: []FmtToken{
		{TokLiteral, 0, "skip;semi foo"}}}},
	{"h_^m*-", "", Section{Type: TimeFormat, SubType: Time, Tokens: []FmtToken{
		{TokHour, 1, ""},
		{TokSpace, 0, "^"},
		{TokMinute, 1, ""},
		{TokRepeat, 0, "-"},
	}}},
}

func (s *CellFmtSuite) TestTokenizeTime(c *C) {
	for _, test := range timeTokenizeTests {
		fmt.Println(test.input)
		section, rem := tokenizeTime(nil, []byte(test.input))
		c.Assert(section, DeepEquals, test.expected, Commentf("token mismatch input=%q", test.input))
		c.Assert(string(rem), DeepEquals, test.rem, Commentf("remainder mismatch input=%q", test.input))
	}
}

var numTokenizeTests = []struct {
	input    string
	rem      string
	expected []FmtToken
}{
	{"0.#", "", []FmtToken{
		{TokNumInt, 0, "0"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "#"},
	}},
	{"0.00 123", "", []FmtToken{
		{TokNumInt, 0, "0"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "00"},
		{TokLiteral, 0, " 123"},
	}},
	{"#.##E+00", "", []FmtToken{
		{TokNumInt, 0, "#"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "##"},
		{TokNumExp, 0, "E+00"},
	}},
	{"#.## 000/000", "", []FmtToken{
		{TokNumInt, 0, "#"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "##"},
		{TokLiteral, 0, " "},
		{TokNumFracNum, 0, "000"},
		{TokNumFracSign, 0, ""},
		{TokNumFracDenom, 0, "000"},
	}},
	{"foo 0.# bar", "", []FmtToken{
		{TokLiteral, 0, "foo "},
		{TokNumInt, 0, "0"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "#"},
		{TokLiteral, 0, " bar"},
	}},
	{"#.##E+00", "", []FmtToken{
		{TokNumInt, 0, "#"},
		{TokNumDecSign, 0, ""},
		{TokNumDec, 0, "##"},
		{TokNumExp, 0, "E+00"},
	}},
	{"# #/16", "", []FmtToken{
		{TokNumInt, 0, "#"},
		{TokLiteral, 0, " "},
		{TokNumFracNum, 0, "#"},
		{TokNumFracSign, 0, ""},
		{TokNumFracDenom, 0, "16"},
	}},
	//{"#.## 00/000", "", []FmtToken{{TokNumInt, 0, "#"}, {TokNumDecSign, 0, ""}, {TokNumDec, 0, "##"}, {TokLiteral, 0, " "}, {TokNumFracNum, 0, "00"}, {TokNumFracSign, 0, ""}, {TokNumFracDenom, 0, "000"}}},
}

func (s *CellFmtSuite) TestTokenizeNumeric(c *C) {
	for _, test := range numTokenizeTests {
		fmt.Println(test.input)
		section, rem := tokenizeNumeric(nil, []byte(test.input))
		expected := Section{Type: NumberFormat, SubType: NoSubType, Tokens: test.expected}
		c.Assert(section, DeepEquals, expected, Commentf("token mismatch input=%q", test.input))
		c.Assert(string(rem), DeepEquals, test.rem, Commentf("remainder mismatch input=%q", test.input))
	}
}

var tokenizeTests = []struct {
	input    string
	expected CellFormat
}{
	{"foo", CellFormat{
		Sections: []Section{
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "foo"}}},
		}}},
	{"h", CellFormat{
		Sections: []Section{
			{TimeFormat, Time, []FmtToken{{TokHour, 1, ""}}},
		}}},
	{"#", CellFormat{
		Sections: []Section{
			{NumberFormat, NoSubType, []FmtToken{{TokNumInt, 0, "#"}}},
		}}},
	{"[h]", CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokTotalHours, 1, ""}}},
		}}},
	{"hh:[z]", CellFormat{
		Sections: []Section{
			{TimeFormat, Time, []FmtToken{{TokHour, 2, ""}, {TokLiteral, 0, ":[z]"}}},
		}}},
	{"hh:[z", CellFormat{
		Sections: []Section{
			{TimeFormat, Time, []FmtToken{{TokHour, 2, ""}, {TokLiteral, 0, ":[z"}}},
		}}},
	{"foo [h]", CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokLiteral, 0, "foo "}, {TokTotalHours, 1, ""}}}}}},
	{"[red][h]", CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokColor, 0, "red"}, {TokTotalHours, 1, ""}}},
		}}},
	{"[=50][h]", CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokCondition, 0, "=50"}, {TokTotalHours, 1, ""}}},
		},
		IsConditional: true,
	}},
	{`[h];"m;s";;text`, CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokTotalHours, 1, ""}}},
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "m;s"}}},
			{TextFormat, NoSubType, nil},
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "text"}}},
		}}},
	{"[h];m;s;text", CellFormat{
		Sections: []Section{
			{TimeFormat, Duration, []FmtToken{{TokTotalHours, 1, ""}}},
			{TimeFormat, Date, []FmtToken{{TokMonth, 1, ""}}},
			{TimeFormat, Time, []FmtToken{{TokSecond, 1, ""}}},
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "text"}}},
		}}},
	{`$general"foo`, CellFormat{
		Sections: []Section{
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "$"}, {TokGeneral, 0, ""}, {TokLiteral, 0, "foo"}}},
		}}},
	{"gz", CellFormat{ // not legal, but make sure it's not treated as "general"
		Sections: []Section{
			{TextFormat, NoSubType, []FmtToken{{TokLiteral, 0, "gz"}}},
		}}},
}

func (s *CellFmtSuite) TestTokenizeFormat(c *C) {
	for _, test := range tokenizeTests {
		toks := ParseFormat(test.input)
		c.Assert(toks, DeepEquals, test.expected, Commentf("input=%q", test.input))
	}
}

var formatValueTests = []struct {
	inval    string
	cellType CellType
	infmt    string
	expected string
}{
	// time
	{"0.007", CellTypeNumeric, "hh:mm:ss.00", "00:10:04.80"},
	{"42099.655960", CellTypeNumeric, "yyyy-mm-dd hh:mm:ss.000 am/pm", "2015-04-05 03:44:34.944 pm"},
	{"42099.655960", CellTypeNumeric, "yyyy\\/mm\\/dd hh:mm:ss am/pm", "2015/04/05 03:44:35 pm"},
	{"42099.655960", CellTypeNumeric, "ddd mmmmm yy", "Sun A 15"},
	{"42099.655960", CellTypeNumeric, "dddd mmmm", "Sunday April"},
	{"42099.655960", CellTypeNumeric, "m/d/yyy", "4/5/2015"},
	{"42099.655960", CellTypeNumeric, "h:m:s AM/PM", "3:44:35 PM"},
	{"42099.655960", CellTypeNumeric, "h:m:s A/P", "3:44:35 P"},

	// duration
	{"2.23802615740741", CellTypeNumeric, "[hh]:mm:ss.00", "53:42:45.46"},
	{"2.23802615740741", CellTypeNumeric, "[hhhh]:mm:ss.00", "0053:42:45.46"},
	{"2.23802615740741", CellTypeNumeric, "[mm]:ss.00", "3222:45.46"},
	{"2.23802615740741", CellTypeNumeric, "[ss].00", "193365.46"},

	// number
	{"1234.56", CellTypeNumeric, "#", "1235"}, // round up
	{"1234.56", CellTypeNumeric, "#,###", "1,235"},
	{"-1234.56", CellTypeNumeric, "#,###", "-1,235"},
	{"123456.78", CellTypeNumeric, "#,", "123"}, // trailing comma should devide by 1000
	{"12345678", CellTypeNumeric, "#,,", "12"},  // trailing comma should devide by 1000
	{"12345678", CellTypeNumeric, "#,###", "12,345,678"},
	{"1", CellTypeNumeric, "#", "1"},
	{"-1", CellTypeNumeric, "#", "-1"},
	{"12", CellTypeNumeric, "0000", "0012"},
	{"1.2", CellTypeNumeric, "#.#", "1.2"},
	{"1.26", CellTypeNumeric, "#.#", "1.3"}, // should round
	{"-1.2", CellTypeNumeric, "#.#", "-1.2"},
	{"-1.2", CellTypeNumeric, "", "-1.2"},
	{"1.2", CellTypeNumeric, "#.##", "1.2"},
	{"1.2", CellTypeNumeric, "#.00", "1.20"},
	{"1.23", CellTypeNumeric, "#.##", "1.23"},
	{"1.234", CellTypeNumeric, "#.##", "1.23"},
	{"1.2", CellTypeNumeric, "#.#0", "1.20"},
	{"1.2", CellTypeNumeric, "?#.#0", " 1.20"},
	{"12.2", CellTypeNumeric, "?#.#0", "12.20"},
	{"1.2", CellTypeNumeric, "#.#?", "1.2 "},
	{"1.23", CellTypeNumeric, "#.#?", "1.23"},

	// zeroes are elided for # in integer formatting
	{"0", CellTypeNumeric, "#", ""},
	{"0", CellTypeNumeric, "#.#", "."},
	{"0", CellTypeNumeric, "#.0", ".0"},
	{"0", CellTypeNumeric, "0.0", "0.0"},

	// comma division
	{"12345678.98", CellTypeNumeric, "#,###.#0,", "12,345.68"}, // trailing comma should devide by 1000

	// percentage
	{"12", CellTypeNumeric, "#%", "1200%"},
	{"1.23", CellTypeNumeric, "0.00%", "123.00%"},
	{"12", CellTypeNumeric, "#,###%", "1,200%"},
	{"0.2345", CellTypeNumeric, "0.####%", "23.45%"},

	// exponents
	{"12345678", CellTypeNumeric, "#E+000", "1E+007"},
	{"0", CellTypeNumeric, "#E+#", "0E+0"}, // must keep the leading zero, even though it begins with #
	{"0", CellTypeNumeric, "#E+00", "0E+00"},
	{"12345678", CellTypeNumeric, "#E+00", "1E+07"},
	{"12345678", CellTypeNumeric, "#E+000", "1E+007"},
	{"12345678", CellTypeNumeric, "#E+##", "1E+7"},
	{"12345678", CellTypeNumeric, "#.##E+00", "1.23E+07"},
	{"12345678", CellTypeNumeric, "#.##e+00", "1.23E+07"},
	{"12345678", CellTypeNumeric, "#.###E+00", "1.235E+07"},

	// fractions
	{"0.75", CellTypeNumeric, "#/#", "3/4"},
	{"0.75", CellTypeNumeric, "#/###", "3/4"},
	{"0.75", CellTypeNumeric, "#/00#", "3/004"},
	{"0.75", CellTypeNumeric, "0#/00#", "03/004"},
	{"0.75", CellTypeNumeric, "?#/00#", " 3/004"},
	{"123.75", CellTypeNumeric, "#/#", "495/4"},
	{"123.75", CellTypeNumeric, "# #/#", "123 3/4"},
	{"-123.75", CellTypeNumeric, "# #/#", "-123 3/4"},
	{"-0.75", CellTypeNumeric, "#/#", "-3/4"},
	{"-0.256", CellTypeNumeric, "# #/#%", "-25 3/5%"},
	{"-0.256", CellTypeNumeric, "# #/$#%", "-25 3/$5%"}, // yes, Excel allows this kind of interleaving
	{"0.25", CellTypeNumeric, "#/16", "4/16"},
	{"0.25", CellTypeNumeric, "#/$16", "4/$16"},

	// general
	{"-1.2", CellTypeNumeric, "general", "-1.2"},
	{"foo", CellTypeString, "@", "foo"},

	// two section tests
	{"1.2", CellTypeNumeric, "#.#;(#.#)", "1.2"},
	{"0", CellTypeNumeric, "#.#;(#.#)", "."}, // zero handled by first section
	{"-1.2", CellTypeNumeric, "#.#;(#.#)", "(1.2)"},
	{"text", CellTypeString, "#.#;(#.#)", "text"},
	{"1.2", CellTypeNumeric, "general;general", "1.2"},
	{"-1.2", CellTypeNumeric, "general", "-1.2"},
	{"-1.2", CellTypeNumeric, "general;general", "1.2"}, // strips the minus sign

	// three section tests
	{"1.2", CellTypeNumeric, "#.#;(#.#);\"iszero\"", "1.2"},
	{"-1.2", CellTypeNumeric, "#.#;(#.#);\"iszero\"", "(1.2)"},
	{"0", CellTypeNumeric, "#.#;(#.#);\"iszero\"", "iszero"},
	{"text", CellTypeString, "#.#;(#.#);\"iszero\"", "text"},

	// four section tests
	{"1.2", CellTypeNumeric, "#.#;(#.#);\"iszero\";\"text >\"@\"< here\"", "1.2"},
	{"-1.2", CellTypeNumeric, "#.#;(#.#);\"iszero\";\"text >\"@\"< here\"", "(1.2)"},
	{"0", CellTypeNumeric, "#.#;(#.#);\"iszero\";\"text >\"@\"< here\"", "iszero"},
	{"text", CellTypeString, "#.#;(#.#);\"iszero\";\"text >\"@\"< here\"", "text >text< here"},
}

func (s *CellFmtSuite) TestFormatValue(c *C) {
	for _, test := range formatValueTests {
		//cell := Cell{Value: test.inval, numFmt: test.infmt, cellType: CellTypeNumeric}
		//f, err := cell.Format()
		f := ParseFormat(test.infmt)
		s, err := f.FormatValue(test.inval, test.cellType, false)
		fmt.Printf("inval=%q infmt=%q output=%q\n", test.inval, test.infmt, s)

		c.Assert(err, IsNil, Commentf("input=%q infmt=%q", test.inval, test.infmt))
		c.Assert(s.FormattedValue, Equals, test.expected, Commentf("input=%q infmt=%q", test.inval, test.infmt))
	}
}

var formattedValueTests = []struct {
	inval    string
	cellType CellType
	infmt    string
	expected FormattedValue
}{
	{"foo", CellTypeString, "", FormattedValue{
		GoValue:        "foo",
		FormattedValue: "foo",
		Section: Section{
			Type:    TextFormat,
			SubType: NoSubType,
		},
	}},
	{"1", CellTypeBool, "", FormattedValue{
		GoValue:        true,
		FormattedValue: "TRUE",
		Section: Section{
			Type:    BoolFormat,
			SubType: NoSubType,
		},
	}},
	{"#VALUE", CellTypeError, "", FormattedValue{
		GoValue:        "#VALUE",
		FormattedValue: "#VALUE",
		Section: Section{
			Type:    ErrorFormat,
			SubType: NoSubType,
		},
	}},
	{"1.26", CellTypeNumeric, "0.#", FormattedValue{
		GoValue:        1.26,
		FormattedValue: "1.3",
		Section: Section{
			Type:    NumberFormat,
			SubType: NoSubType,
			Tokens: []FmtToken{
				{TokNumInt, 0, "0"},
				{TokNumDecSign, 0, ""},
				{TokNumDec, 0, "#"},
			},
		},
	}},
	{"-1.26", CellTypeNumeric, "0.#;(00.##)", FormattedValue{
		GoValue:        -1.26,
		FormattedValue: "(01.26)",
		Section: Section{
			Type:    NumberFormat,
			SubType: NoSubType,
			Tokens: []FmtToken{
				{TokLiteral, 0, "("},
				{TokNumInt, 0, "00"},
				{TokNumDecSign, 0, ""},
				{TokNumDec, 0, "##"},
				{TokLiteral, 0, ")"},
			},
		},
	}},
	{"2.5", CellTypeNumeric, "[hh]:mm", FormattedValue{
		GoValue:        60 * time.Hour,
		FormattedValue: "60:00",
		Section: Section{
			Type:    TimeFormat,
			SubType: Duration,
			Tokens: []FmtToken{
				{TokTotalHours, 2, ""},
				{TokLiteral, 0, ":"},
				{TokMinute, 2, ""},
			},
		},
	}},
	{"42099.625", CellTypeNumeric, "yyyy-mm-dd", FormattedValue{
		GoValue:        time.Date(2015, 4, 5, 15, 0, 0, 0, time.UTC),
		FormattedValue: "2015-04-05",
		Section: Section{
			Type:    TimeFormat,
			SubType: Date,
			Tokens: []FmtToken{
				{TokYear, 4, ""},
				{TokLiteral, 0, "-"},
				{TokMonth, 2, ""},
				{TokLiteral, 0, "-"},
				{TokDay, 2, ""},
			},
		},
	}},
	{"42099.625", CellTypeNumeric, "hh:mm", FormattedValue{
		GoValue:        time.Date(2015, 4, 5, 15, 0, 0, 0, time.UTC),
		FormattedValue: "15:00",
		Section: Section{
			Type:    TimeFormat,
			SubType: Time,
			Tokens: []FmtToken{
				{TokHour, 2, ""},
				{TokLiteral, 0, ":"},
				{TokMinute, 2, ""},
			},
		},
	}},
	{"42099.625", CellTypeNumeric, "yyyy-mm-dd hh:mm", FormattedValue{
		GoValue:        time.Date(2015, 4, 5, 15, 0, 0, 0, time.UTC),
		FormattedValue: "2015-04-05 15:00",
		Section: Section{
			Type:    TimeFormat,
			SubType: DateTime,
			Tokens: []FmtToken{
				{TokYear, 4, ""},
				{TokLiteral, 0, "-"},
				{TokMonth, 2, ""},
				{TokLiteral, 0, "-"},
				{TokDay, 2, ""},
				{TokLiteral, 0, " "},
				{TokHour, 2, ""},
				{TokLiteral, 0, ":"},
				{TokMinute, 2, ""},
			},
		},
	}},
}

func (s *CellFmtSuite) TestFmtFormattedValue(c *C) {
	for _, test := range formattedValueTests {
		f := ParseFormat(test.infmt)
		fv, err := f.FormatValue(test.inval, test.cellType, false)
		c.Assert(err, IsNil, Commentf("input=%q infmt=%q", test.inval, test.infmt))
		c.Assert(fv, DeepEquals, test.expected, Commentf("input=%q infmt=%q", test.inval, test.infmt))
	}
}

var thouTests = []struct {
	in       string
	expected string
}{
	{"", ""},
	{"1", "1"},
	{"12", "12"},
	{"123", "123"},
	{"1234", "1,234"},
	{"123456", "123,456"},
	{"1234567", "1,234,567"},
}

func (s *CellFmtSuite) TestThou(c *C) {
	for _, test := range thouTests {
		out := fmtThou(test.in)

		c.Assert(out, Equals, test.expected, Commentf("input=%q", test.in))
	}
}
