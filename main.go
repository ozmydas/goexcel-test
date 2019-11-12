package main

import (
    "fmt"

    "github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
    // baris := make(map[int]interface{})
    // kolom := make(map[string]interface{})

    f, err := excelize.OpenFile("./tes.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }

    // Get all t
    ir := 1
    var x string
    rows, err := f.GetRows("Sheet1")
    for _, row := range rows {
        ic := 0
        for _, colCell := range row {

            ic++
            if colCell == "" {
                continue
            }

            x = toCharStrArr(ic)
            fmt.Printf("%d%s : %s\n", ir, x, colCell)
        }
        fmt.Println("==========")
        ir++
    }

}

var arr = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF"}

func toCharStrArr(i int) string {
    return arr[i-1]
}
