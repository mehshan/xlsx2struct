package xlsx2struct

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestParseTag(t *testing.T) {
	tag := parseColumnTag("heading=Order Date,trim,time=2006-01-02,default=None")
	require.NotNil(t, tag)
	// TODO: complete test cases
}
