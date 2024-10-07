package controllers

import (
	"bytes"
	"fmt"
	"log"
	"net/http"
	"project-its/initializers"
	"project-its/models"
	"strconv"
	"time"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

// Create a new event
func GetEventsRapat(c *gin.Context) {
	var events []models.JadwalRapat
	if err := initializers.DB.Find(&events).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	c.JSON(http.StatusOK, gin.H{"rapat": events})
}

func CreateEventRapat(c *gin.Context) {
	var event models.JadwalRapat
	if err := c.ShouldBindJSON(&event); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	// Set notification menggunakan fungsi dari notificationController
	loc, err := time.LoadLocation("Asia/Jakarta")
	if err != nil {
		log.Printf("Error loading location: %v", err)
		return
	}

	var startTime time.Time
	if event.AllDay {
		// Jika AllDay = true, set waktu ke tengah malam
		startTime, err = time.ParseInLocation("2006-01-02T15:04:05", event.Start+"T00:00:00", loc)
	} else {
		// Jika tidak, parse dengan format RFC3339
		startTime, err = time.ParseInLocation(time.RFC3339, event.Start, loc)
	}

	if err != nil {
		log.Printf("Error parsing start time: %v", err)
		return
	}

	SetNotification(event.Title, startTime, "JadwalRapat") // Panggil fungsi SetNotification

	if err := initializers.DB.Create(&event).Error; err != nil {
		log.Printf("Error creating event: %v", err)
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	c.JSON(http.StatusOK, event)
}

func DeleteEventRapat(c *gin.Context) {
	id := c.Param("id") // Menggunakan c.Param jika UUID dikirim sebagai bagian dari URL
	if id == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "ID harus disertakan"})
		return
	}
	if err := initializers.DB.Where("id = ?", id).Delete(&models.JadwalRapat{}).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	c.Status(http.StatusNoContent)
}

func ExportJadwalRapatToExcel(c *gin.Context) {
	// Ambil data dari model JadwalRapat
	var events_rapat []models.JadwalRapat
	if err := initializers.DB.Find(&events_rapat).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	f := excelize.NewFile()
	sheet := "Calendar 2024"
	f.NewSheet(sheet)

	months := []string{
		"January 2024", "February 2024", "March 2024", "April 2024",
		"May 2024", "June 2024", "July 2024", "August 2024",
		"September 2024", "October 2024", "November 2024", "December 2024",
	}

	rowOffset := 0
	colOffset := 0
	for i, month := range months {
		setMonthDataRapat(f, sheet, month, rowOffset, colOffset, events_rapat)
		colOffset += 9 // Sesuaikan offset untuk bulan berikutnya dalam baris yang sama
		if (i+1)%3 == 0 {
			rowOffset += 18 // Pindah ke baris berikutnya setiap 3 bulan
			colOffset = 0
		}
	}

	// Hapus sheet default
	f.DeleteSheet("Sheet1")

	// Simpan file ke buffer
	var buffer bytes.Buffer
	if err := f.Write(&buffer); err != nil {
		fmt.Println(err)
		return
	}

	// Set header untuk download file
	c.Header("Content-Disposition", "attachment; filename=Calendar2024.xlsx")
	c.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	c.Header("Content-Length", strconv.Itoa(len(buffer.Bytes())))
	c.Writer.Write(buffer.Bytes())
}

func setMonthDataRapat(f *excelize.File, sheet, month string, rowOffset, colOffset int, events_rapat []models.JadwalRapat) {
	var (
		monthStyle, titleStyle, dataStyle, blankStyle,
		grayBlankStyle, grayDataStyle int
		err  error
		addr string
	)
	// Get the first day of the month and the number of days in the month
	monthTime, err := time.Parse("January 2006", month)
	if err != nil {
		fmt.Println(err)
		return
	}
	firstDay := monthTime.Weekday()
	daysInMonth := time.Date(monthTime.Year(), monthTime.Month()+1, 0, 0, 0, 0, 0, time.UTC).Day()

	// cell values
	data := map[int][]interface{}{
		1 + rowOffset: {month},
		3 + rowOffset: {"SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY",
			"THURSDAY", "FRIDAY", "SATURDAY"},
	}

	// Fill in the dates
	day := 1
	for r := 4 + rowOffset; day <= daysInMonth; r += 2 {
		week := make([]interface{}, 7)
		eventDetails := make([]interface{}, 7)
		for d := firstDay; d < 7 && day <= daysInMonth; d++ {
			week[d] = day

			// Cek apakah ada event pada hari ini
			for _, event := range events_rapat {
				startDate, _ := time.Parse(time.RFC3339, event.Start)
				endDate, _ := time.Parse(time.RFC3339, event.End)
				currentDate := time.Date(monthTime.Year(), monthTime.Month(), day, 0, 0, 0, 0, time.UTC)

				if (currentDate.Equal(startDate) || currentDate.After(startDate)) && currentDate.Before(endDate.AddDate(0, 0, 1)) {
					startTime := startDate.Format("15:04")
					endTime := endDate.Format("15:04")
					eventDetails[d] = fmt.Sprintf("%s\n%s - %s", event.Title, startTime, endTime)
				}
			}

			day++
		}
		data[r] = week
		data[r+1] = eventDetails
		firstDay = 0 // Reset firstDay for subsequent weeks
	}

	// custom rows height
	height := map[int]float64{
		1 + rowOffset: 45, 3 + rowOffset: 22, 5 + rowOffset: 30, 7 + rowOffset: 30,
		9 + rowOffset: 30, 11 + rowOffset: 30, 13 + rowOffset: 30, 15 + rowOffset: 30,
	}
	top := excelize.Border{Type: "top", Style: 1, Color: "DADEE0"}
	left := excelize.Border{Type: "left", Style: 1, Color: "DADEE0"}
	right := excelize.Border{Type: "right", Style: 1, Color: "DADEE0"}
	bottom := excelize.Border{Type: "bottom", Style: 1, Color: "DADEE0"}
	fill := excelize.Fill{Type: "pattern", Color: []string{"EFEFEF"}, Pattern: 1}

	// set each cell value
	for r, row := range data {
		if addr, err = excelize.JoinCellName(string('B'+colOffset), r); err != nil {
			fmt.Println(err)
			return
		}
		if err = f.SetSheetRow(sheet, addr, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// set custom row height
	for r, ht := range height {
		if err = f.SetRowHeight(sheet, r, ht); err != nil {
			fmt.Println(err)
			return
		}
	}
	// set custom column width
	if err = f.SetColWidth(sheet, string('B'+colOffset), string('H'+colOffset), 10); err != nil {
		fmt.Println(err)
		return
	}
	// merge cell for the 'MONTH'
	if err = f.MergeCell(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), 1+rowOffset), fmt.Sprintf("%s%d", string('D'+colOffset), 1+rowOffset)); err != nil {
		fmt.Println(err)
		return
	}
	// define font style for the 'MONTH'
	if monthStyle, err = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "1f7f3b", Bold: true, Size: 22, Family: "Arial"},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set font style for the 'MONTH'
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), 1+rowOffset), fmt.Sprintf("%s%d", string('D'+colOffset), 1+rowOffset), monthStyle); err != nil {
		fmt.Println(err)
		return
	}
	// define style for the 'SUNDAY' to 'SATURDAY'
	if titleStyle, err = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "1f7f3b", Size: 10, Bold: true, Family: "Arial"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E6F4EA"}, Pattern: 1},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Border:    []excelize.Border{{Type: "top", Style: 2, Color: "1f7f3b"}},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set style for the 'SUNDAY' to 'SATURDAY'
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), 3+rowOffset), fmt.Sprintf("%s%d", string('H'+colOffset), 3+rowOffset), titleStyle); err != nil {
		fmt.Println(err)
		return
	}
	// define cell border for the date cell in the date range
	if dataStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{top, left, right},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set cell border for the date cell in the date range
	for _, r := range []int{4, 6, 8, 10, 12, 14} {
		if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), r+rowOffset),
			fmt.Sprintf("%s%d", string('H'+colOffset), r+rowOffset), dataStyle); err != nil {
			fmt.Println(err)
			return
		}
	}
	// define cell border for the blank cell in the date range
	if blankStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{left, right, bottom},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set cell border for the blank cell in the date range
	for _, r := range []int{5, 7, 9, 11, 13, 15} {
		if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), r+rowOffset),
			fmt.Sprintf("%s%d", string('H'+colOffset), r+rowOffset), blankStyle); err != nil {
			fmt.Println(err)
			return
		}
	}
	// define the border and fill style for the blank cell in previous and next month
	if grayBlankStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{left, right, bottom},
		Fill:   fill}); err != nil {
		fmt.Println(err)
		return
	}
	// set the border and fill style for the blank cell in previous and next month
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), 5+rowOffset), fmt.Sprintf("%s%d", string('F'+colOffset), 5+rowOffset), grayBlankStyle); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('C'+colOffset), 15+rowOffset), fmt.Sprintf("%s%d", string('H'+colOffset), 15+rowOffset), grayBlankStyle); err != nil {
		fmt.Println(err)
		return
	}
	// define the border and fill style for the date cell in previous and next month
	if grayDataStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{left, right, top},
		Font:   &excelize.Font{Color: "777777"}, Fill: fill}); err != nil {
		fmt.Println(err)
		return
	}
	// set the border and fill style for the date cell in previous and next month
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('B'+colOffset), 4+rowOffset), fmt.Sprintf("%s%d", string('F'+colOffset), 4+rowOffset), grayDataStyle); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", string('C'+colOffset), 14+rowOffset), fmt.Sprintf("%s%d", string('H'+colOffset), 14+rowOffset), grayDataStyle); err != nil {
		fmt.Println(err)
		return
	}
	// hide gridlines for the worksheet
	disable := false
	if err := f.SetSheetView(sheet, 0, &excelize.ViewOptions{
		ShowGridLines: &disable,
	}); err != nil {
		fmt.Println(err)
	}
}