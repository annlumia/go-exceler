package exceler

import "regexp"

var (
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rangeRgx    = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
)
