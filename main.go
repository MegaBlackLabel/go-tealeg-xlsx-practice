package main

import (
	"errors"
	"fmt"

	"github.com/tealeg/xlsx/v3"
)

func cellVisitor(c *xlsx.Cell) error {
	x, y := c.GetCoordinates()
	fmt.Printf("x: %d, y: %d ::: ", x, y)
	value, err := c.FormattedValue()
	if err != nil {
		fmt.Println(err.Error())
	} else {
		fmt.Println("Cell value:", value)
	}
	return err
}

func rowVisitor(r *xlsx.Row) error {
	num := r.GetCoordinate()
	fmt.Println("Cell num:", num)
	return r.ForEachCell(cellVisitor)
}

func rowStuff() {
	filename := "samplefile.xlsx"
	wb, err := xlsx.OpenFile(filename)
	if err != nil {
		panic(err)
	}
	sh, ok := wb.Sheet["Sample"]
	if !ok {
		panic(errors.New("Sheet not found"))
	}
	fmt.Println("Max row is", sh.MaxRow)
	sh.ForEachRow(rowVisitor)
}

func main() {
	// open an existing file
	wb, err := xlsx.OpenFile("samplefile.xlsx")
	if err != nil {
		panic(err)
	}
	// wb now contains a reference to the workbook
	// show all the sheets in the workbook
	fmt.Println("Sheets in this file:")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("----")

	sh := wb.Sheet["Styles"]
	cell, err := sh.Cell(0, 1)
	if err != nil {
		panic(err)
	}
	style := cell.GetStyle()
	fmt.Println("Cell value:", cell.String())
	fmt.Println("Font:", style.Font.Name)
	fmt.Println("Size:", style.Font.Size)
	fmt.Println("H-Align:", style.Alignment.Horizontal)
	fmt.Println("ForeColor:", style.Fill.FgColor)
	fmt.Println("BackColor:", style.Fill.BgColor)
	fmt.Println("----")

	// get the Cell in D1, which is row 0, col 3
	sh2 := wb.Sheet["Sample"]
	theCell, err := sh2.Cell(0, 3)
	if err != nil {
		panic(err)
	}
	// we got a cell, but what's in it?
	fv, err := theCell.FormattedValue()
	if err != nil {
		panic(err)
	}
	fmt.Println("Numeric cell?:", theCell.Type() == xlsx.CellTypeNumeric)
	fmt.Println("String:", theCell.String())
	fmt.Println("Formatted:", fv)
	fmt.Println("Formula:", theCell.Formula())

	fmt.Println("== xlsx package tutorial ==")
	rowStuff()
}
