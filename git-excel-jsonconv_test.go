package main

import (
	"os"
	"testing"
)

func Test_xls(t *testing.T) {
	os.Args = []string{"", "test.xls"}

	main()
}

func Test_xlsx(t *testing.T) {
	os.Args = []string{"", "test.xlsx"}

	main()
}
