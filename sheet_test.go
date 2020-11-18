package streamxlsx

import (
	"testing"
)

func TestAsRef(t *testing.T) {
	mustEq(t, "A1", AsRef(0, 0))
	mustEq(t, "B1", AsRef(1, 0))
	mustEq(t, "C1", AsRef(2, 0))
	mustEq(t, "A2", AsRef(0, 1))
	mustEq(t, "B2", AsRef(1, 1))
	mustEq(t, "C2", AsRef(2, 1))
	mustEq(t, "Z1", AsRef(25, 0))
	mustEq(t, "AA1", AsRef(1*26, 0))
	mustEq(t, "AZ1", AsRef(1*26+25, 0))
	mustEq(t, "BA1", AsRef(2*26, 0))
	mustEq(t, "ZA1", AsRef(26*26, 0))
	mustEq(t, "ZZ1", AsRef(26*26+25, 0))

	mustEq(t, "A9", AsRef(0, 8))
	mustEq(t, "A10", AsRef(0, 9))
	mustEq(t, "AA10", AsRef(26, 9))

	mustEq(t, "AAA1", AsRef(27*26, 0))
	mustEq(t, "AAB1", AsRef(27*26+1, 0))
	mustEq(t, "AAZ1", AsRef(27*26+25, 0))
	mustEq(t, "ABA1", AsRef(27*26+26, 0))
}

func mustEq(t *testing.T, want, have string) {
	t.Helper()
	if have != want {
		t.Fatalf("have %q, want %q", have, want)
	}
}
