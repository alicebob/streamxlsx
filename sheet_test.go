package streamxlsx

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAsRef(t *testing.T) {
	assert.Equal(t, "A1", AsRef(0, 0))
	assert.Equal(t, "B1", AsRef(1, 0))
	assert.Equal(t, "C1", AsRef(2, 0))
	assert.Equal(t, "A2", AsRef(0, 1))
	assert.Equal(t, "B2", AsRef(1, 1))
	assert.Equal(t, "C2", AsRef(2, 1))
	assert.Equal(t, "Z1", AsRef(25, 0))
	assert.Equal(t, "AA1", AsRef(1*26, 0))
	assert.Equal(t, "AZ1", AsRef(1*26+25, 0))
	assert.Equal(t, "BA1", AsRef(2*26, 0))
	assert.Equal(t, "ZA1", AsRef(26*26, 0))
	assert.Equal(t, "ZZ1", AsRef(26*26+25, 0))

	assert.Equal(t, "A9", AsRef(0, 8))
	assert.Equal(t, "A10", AsRef(0, 9))
	assert.Equal(t, "AA10", AsRef(26, 9))

	assert.Equal(t, "AAA1", AsRef(27*26, 0))
	assert.Equal(t, "AAB1", AsRef(27*26+1, 0))
	assert.Equal(t, "AAZ1", AsRef(27*26+25, 0))
	assert.Equal(t, "ABA1", AsRef(27*26+26, 0))
}
