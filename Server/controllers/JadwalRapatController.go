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
	c.Header("Content-Disposition", "attachment; filename=its_report_jadwalRapat.xlsx")
	c.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	c.Header("Content-Length", strconv.Itoa(len(buffer.Bytes())))
	c.Writer.Write(buffer.Bytes())
}

func setMonthDataRapat(f *excelize.File, sheet, month string, rowOffset, colOffset int, events_rapat []models.JadwalRapat) {
	var (
		monthStyle, titleStyle, dataStyle, blankStyle int
		err                                           error
		addr                                          string
	)

	// Definisikan loc di sini
	loc, err := time.LoadLocation("Asia/Jakarta")
	if err != nil {
		fmt.Printf("Error loading location: %v", err)
		return
	}

	monthTime, err := time.Parse("January 2006", month)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get the first day of the month and the number of days in the month
	firstDay := monthTime.Weekday()
	daysInMonth := time.Date(monthTime.Year(), monthTime.Month()+1, 0, 0, 0, 0, 0, loc).Day()

	// cell values
	data := map[int][]interface{}{
		1 + rowOffset: {month},
		3 + rowOffset: {"MINGGU", "SENIN", "SELASA", "RABU",
			"KAMIS", "JUMAT", "SABTU"},
	}

	// Fill in the dates
	day := 1
	for r := 4 + rowOffset; day <= daysInMonth; r += 2 {
		week := make([]interface{}, 7) // Inisialisasi ulang array week untuk setiap baris baru
		eventDetails := make([]interface{}, 7)
		for d := 0; d < 7; d++ { // Mulai loop dari 0 hingga 6 (Minggu hingga Sabtu)
			if r == 4 + rowOffset && d < int(firstDay) {
				// Jika ini adalah baris pertama dan hari ini sebelum 'firstDay', biarkan kosong
				continue
			}
			if day <= daysInMonth {
				week[d] = day // Isi tanggal

				// Cek apakah ada event pada hari ini
				for _, event := range events_rapat {
					var startDate, endDate time.Time
					if event.AllDay {
						startDate, err = time.ParseInLocation("2006-01-02", event.Start[:10], loc)
						if err != nil {
							fmt.Printf("Error parsing all-day start date: %v\n", err)
							continue
						}
						endDate = startDate // Set endDate to startDate for all-day events
					} else {
						startDate, _ = time.Parse("2006-01-02T15:04:05", event.Start[:10]+"T00:00:00")
						endDate, _ = time.Parse("2006-01-02T15:04:05", event.End[:10]+"T00:00:00")
					}
					currentDate := time.Date(monthTime.Year(), monthTime.Month(), day, 0, 0, 0, 0, time.UTC) // Pastikan waktu diatur ke 00:00:00
					fmt.Printf("Current Date: %s, Start Date: %s, End Date: %s\n", currentDate, startDate, endDate)

					// Periksa apakah currentDate sama dengan startDate atau berada di antara startDate dan endDate
					if currentDate.Equal(startDate) || (currentDate.After(startDate) && currentDate.Before(endDate.AddDate(0, 0, 1))) {
						var eventDetail string
						if event.AllDay {
							eventDetail = fmt.Sprintf("%s\nAllDay", event.Title)
						} else {
							startDate, _ = time.Parse(time.RFC3339, event.Start)
							endDate, _ = time.Parse(time.RFC3339, event.End)
							// Format hanya jam dan menit dari tanggal
							eventDetail = fmt.Sprintf("%s\n%s - %s", event.Title, startDate.Format("15:04"), endDate.Format("15:04"))
						}

						// Gabungkan detail acara jika sudah ada
						if eventDetails[d] != nil {
							eventDetails[d] = fmt.Sprintf("%s\n%s", eventDetails[d], eventDetail)
						} else {
							eventDetails[d] = eventDetail
						}
						// Hanya terapkan warna untuk event pertama pada hari itu
						if event.Title != "" && eventDetails[d] == eventDetail { // Perubahan kondisi di sini
							cellAddr, _ := excelize.JoinCellName(string('B'+colOffset+int(d)), r+1)

							// Buat gaya baru dengan warna latar belakang
							style, err := f.NewStyle(&excelize.Style{
								Fill: excelize.Fill{
									Type:    "pattern",
									Color:   []string{event.Color},
									Pattern: 1,
								},
								Font:      &excelize.Font{Size: 9},
								Alignment: &excelize.Alignment{WrapText: true},
								Border: []excelize.Border{
									{Type: "left", Color: "DADEE0", Style: 1},
									{Type: "right", Color: "DADEE0", Style: 1},
									{Type: "top", Color: "DADEE0", Style: 1},
									{Type: "bottom", Color: "DADEE0", Style: 1},
								},
							})
							if err != nil {
								fmt.Printf("Error membuat gaya untuk sel %s: %v\n", cellAddr, err)
								continue
							}

							// Terapkan gaya ke sel
							if err := f.SetCellStyle(sheet, cellAddr, cellAddr, style); err != nil {
								fmt.Printf("Error menerapkan gaya ke sel %s: %v\n", cellAddr, err)
							} else {
								fmt.Printf("Berhasil menerapkan gaya dengan warna %s ke sel %s\n", event.Color, cellAddr)
							}
						}
					}
				}

				day++ // Increment day hanya jika hari ini diisi
			}
		}
		data[r] = week
		data[r+1] = eventDetails
		if r == 4 + rowOffset {
			firstDay = 0 // Reset firstDay untuk minggu berikutnya
		}
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
	if err = f.SetColWidth(sheet, string('B'+colOffset), string('H'+colOffset), 15); err != nil {
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
		Border:    []excelize.Border{left, right, bottom},
		Font:      &excelize.Font{Size: 9},
		Alignment: &excelize.Alignment{WrapText: true},
	}); err != nil {
		fmt.Println(err)
		return
	}

	// set cell border for the blank cell in the date range, but only for cells that don't have a fill color
	for _, r := range []int{5, 7, 9, 11, 13, 15} {
		for c := 0; c < 7; c++ {
			cellAddr, _ := excelize.JoinCellName(string('B'+colOffset+c), r+rowOffset)
			if styleID, err := f.GetCellStyle(sheet, cellAddr); err != nil {
				// Handle error
				fmt.Println("Error mendapatkan gaya sel:", err)
				return
			} else if styleID == 0 {
				// Jika tidak ada gaya yang diterapkan, maka terapkan blankStyle
				if err = f.SetCellStyle(sheet, cellAddr, cellAddr, blankStyle); err != nil {
					fmt.Println("Error menerapkan blankStyle:", err)
					return
				}
			}
		}
	}

	// hide gridlines for the worksheet
	disable := false
	if err := f.SetSheetView(sheet, 0, &excelize.ViewOptions{
		ShowGridLines: &disable,
	}); err != nil {
		fmt.Println(err)
	}
}
