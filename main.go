package main

import (
    "fmt"
    "time"

    "github.com/360EntSecGroup-Skylar/excelize"
)

// struct untuk menampung data tiap cell
type cellData struct {
    Cell  string
    Value interface{}
}

func main() {
    // baca dari file tes.xlsx
    f, err := excelize.OpenFile("./tes.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }

    // looping tiap cell yg ada di file
    ir := 1
    fix_ir := 1
    var x, fix_x string
    datax := []cellData{}
    rows, err := f.GetRows("Sheet1")

    // looping baris
    for _, row := range rows {
        ic := 0
        fix_ic := 1
        ada := false

        // looping kolom
        for _, colCell := range row {
            ic++

            // jika cell kosong, langsung skip ke cell selanjutnya
            if colCell == "" {
                continue
            }

            // mengubah index yang tadinya integer jadi huruf, misal 1 = A, 2 = B, dst..
            x = toCharStrArr(ic)
            fix_x = toCharStrArr(fix_ic)

            fmt.Printf("Reading Cell %s%d : %s\n", x, ir, colCell)

            // tampung dan masukkan value dari cell menjadi array
            iden := fmt.Sprint(fix_x, fix_ir)
            data := cellData{Cell: iden, Value: colCell}

            datax = append(datax, data)
            fix_ic++
            ada = true
        }

        ir++

        if ada {
            fix_ir++
        }
    }

    fmt.Println("==========")
    fmt.Println("Let's Remove the Empty Row(s) and Cell(s)")
    fmt.Println("==========")

    // unah array hasil dari yg diatas jadi file excel baru
    create(datax)
} // end func

func create(datax []cellData) {
    f := excelize.NewFile()
    // buat sheet baru
    index := f.NewSheet("Sheet1")

    // set value dari masing2 cell
    for _, elem := range datax {
        cellVal := fmt.Sprintf("%v", elem.Value)
        fmt.Println("Writing Cell " + elem.Cell + " : " + cellVal)
        f.SetCellValue("Sheet1", elem.Cell, cellVal)
    }

    // set sheet1 menjadi default
    f.SetActiveSheet(index)

    // simpan ke file excel baru sesuai tanggal
    timeStr := time.Now().Format("2006-01-02_15:04:05")
    err := f.SaveAs("./output/" + timeStr + ".xlsx")
    if err != nil {
        fmt.Println(err)
    }

    fmt.Println("New Excel Created in Output Folder : " + timeStr + ".xlsx")
} // end func

/******/

var arr = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func toCharStrArr(i int) string {
    if i > 26 {
        puluhan := i / 26
        selisih := i % 26

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
