package main

import (
	"database/sql"
	"fmt"
	_ "odbc/driver"
	"os"
	"time"

	"github.com/go-ini/ini"
	"github.com/xuri/excelize"
)

func main() {
	cfg, err := ini.Load("setting.conf")
	if err != nil {
		fmt.Printf("Fail to read file: %v", err)
		os.Exit(1)
	}

	sql2 := cfg.Section("sql").Key("sql").String()
	host := cfg.Section("database").Key("host").String()
	user := cfg.Section("database").Key("user").String()
	pass := cfg.Section("database").Key("pass").String()
	name := cfg.Section("database").Key("name").String()

	connStr := fmt.Sprintf("driver={SQL Server};SERVER=%s;UID=%s;PWD=%s;DATABASE=%s", host, user, pass, name)
	conn, err := sql.Open("odbc", connStr)
	if err != nil {
		fmt.Println("Connecting Error")
		return
	}
	defer conn.Close()
	stmt, err := conn.Prepare(sql2)
	if err != nil {
		fmt.Println("Query Error", err)
		return
	}
	defer stmt.Close()
	row, err := stmt.Query()
	if err != nil {
		fmt.Println("Query Error", err)
		return
	}
	defer row.Close()

	//获取记录列

	if columns, err := row.Columns(); err != nil {
		fmt.Println("read column Error", err)
		return
	} else {
		//拼接记录Map
		values := make([]sql.RawBytes, len(columns))
		scans := make([]interface{}, len(columns))
		for i := range values {
			scans[i] = &values[i]
		}
		fmt.Println(scans)
		var cells = [26]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
		f := excelize.NewFile()
		index := f.GetActiveSheetIndex()
		idx := 2
		for row.Next() {
			_ = row.Scan(scans...)
			each := map[string]interface{}{}
			for i, col := range values {
				if idx == 2 {
					f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", cells[i], 1), columns[i])
					fmt.Printf("%s%d -> %s\n", cells[i], 1, columns[i])
				}
				each[columns[i]] = string(col)
				fmt.Printf("%s%d -> %s\n", cells[i], idx, col)
				er := f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", cells[i], idx), string(col))
				fmt.Println(er)
			}
			idx++

			fmt.Println(each)
		}
		// // 根据指定路径保存文件
		f.SetActiveSheet(index)
		fileName := time.Now().Format("20060102150405")
		if err := f.SaveAs(fileName + ".xlsx"); err != nil {
			fmt.Println(err)
		}

	}

	fmt.Printf("%s\n", "finish")
	return
}
