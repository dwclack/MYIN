package main

import (
	"fmt"

	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
)

type SS struct {
	user_id       string
	task_id       string
	task_status   string
	task_type     string
	create_time   string
	game_id       string
	server_id     string
	cps_id        string
	reg_server_id string
	reg_cps_id    string
}

// select user_id,task_id,task_status,task_type,create_time,game_id,server_id,cps_id,reg_server_id,reg_cps_id from log_task
//
//	where user_id >= 94800154 and user_id <=94800997
func main() {
	f, err := excelize.OpenFile("2023年4月1日0点~11点创角.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	f2 := excelize.NewFile()

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	var sss = make(map[string][]SS)
	for _, row := range rows {
		var ss = SS{
			user_id:       row[0],
			task_id:       row[1],
			task_status:   row[2],
			task_type:     row[3],
			create_time:   row[4],
			game_id:       row[5],
			server_id:     row[6],
			cps_id:        row[7],
			reg_server_id: row[8],
			reg_cps_id:    row[9],
		}
		shuzu := sss[ss.user_id]
		if shuzu == nil {
			shuzu = []SS{}
		}
		shuzu = append(shuzu, ss)
		sss[ss.user_id] = shuzu
		// for _, colCell := range row {
		// 	fmt.Print(colCell, "\t")
		// }
		// fmt.Println()
	}

	var i = 1
	for _, v := range sss {
		i++
		// 求出最大的时间
		var max SS
		max = v[0]
		for _, v2 := range v {
			if v2.create_time > max.create_time {
				max = v2
			}
		}
		if cast.ToInt32(max.user_id) == 0 {
			continue
		}
		// Set value of a cell.
		f2.SetCellValue("Sheet1", "A"+cast.ToString(i), cast.ToInt32(max.user_id))
		f2.SetCellValue("Sheet1", "B"+cast.ToString(i), cast.ToInt32(max.task_id))
		f2.SetCellValue("Sheet1", "C"+cast.ToString(i), cast.ToInt32(max.task_status))
		f2.SetCellValue("Sheet1", "D"+cast.ToString(i), cast.ToInt32(max.task_type))
		f2.SetCellValue("Sheet1", "E"+cast.ToString(i), max.create_time)
		f2.SetCellValue("Sheet1", "F"+cast.ToString(i), cast.ToInt32(max.game_id))
		f2.SetCellValue("Sheet1", "G"+cast.ToString(i), cast.ToInt32(max.server_id))
		f2.SetCellValue("Sheet1", "H"+cast.ToString(i), cast.ToInt32(max.cps_id))
		f2.SetCellValue("Sheet1", "I"+cast.ToString(i), cast.ToInt32(max.reg_server_id))
		f2.SetCellValue("Sheet1", "J"+cast.ToString(i), cast.ToInt32(max.reg_cps_id))
	}
	// Save spreadsheet by the given path.
	if err := f2.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
