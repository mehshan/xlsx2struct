package xlsx2struct

import (
	"strings"
)

type columnTag struct {
	heading      string
	trim         bool
	timeFormats  []string
	defaultValue string
}

const (
	ColumnTag = "column"

	HeadingOption = "heading"
	TrimOption    = "trim"
	DefaultOption = "default"
	TimeOption    = "time"
)

func parseColumnTag(str string) columnTag {
	t := columnTag{}
	opts := strings.Split(str, ",")

	for _, opt := range opts {
		kv := strings.Split(opt, "=")
		k := strings.ToLower(strings.TrimSpace(kv[0]))

		var v string
		if len(kv) > 1 {
			v = kv[1]
		}

		switch k {
		case HeadingOption:
			t.heading = v
		case TrimOption:
			t.trim = true
		case DefaultOption:
			t.defaultValue = v
		case TimeOption:
			t.timeFormats = []string{v}
		}
	}

	return t
}
