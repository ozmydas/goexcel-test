package main

import (
	"fmt"
	// "strconv"
)

func main() {
	num := 79
	x := toCharStrArr(num)

	fmt.Println(x)
}

// fungsi untuk merubah integer jadi alphabet
var arr = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
	"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func toCharStrArr(i int) string {
	if i > 26 {
		puluhan := i / 26
		selisih := i % 26

		// s_selisih := strconv.Itoa(selisih)
		// s_puluhan := strconv.Itoa(puluhan)
		// fmt.Println(selisih)

		if selisih == 0 {
			puluhan = puluhan - 1
			selisih = 26
		}

		depan := toCharStrArr(puluhan)
		belakang := toCharStrArr(selisih)
		return depan + belakang
	}

	return arr[i-1]
} // end func
