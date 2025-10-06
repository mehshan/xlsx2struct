package xlsx2struct

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestParseTag(t *testing.T) {
	tag := parseColumnTag("heading=Order Date,trim,time=2006-01-02,default=None")
	require.NotNil(t, tag)
	require.Equal(t, "Order Date", tag.heading)
	require.Equal(t, true, tag.trim)
	require.Equal(t, []string{"2006-01-02"}, tag.timeFormats)
	require.Equal(t, "None", tag.defaultValue)
}
