package controllers

import (
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"project-its/initializers"
	"project-its/models"
	"strconv"
	"strings"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

func ExportAllSheets(c *gin.Context) {
	dir := "C:\\excel"
	baseFileName := "its_report"
	filePath := filepath.Join(dir, baseFileName+".xlsx")

	// Check if the file already exists
	if _, err := os.Stat(filePath); err == nil {
		// File exists, append "_new" to the file name
		baseFileName += "_new"
	}

	fileName := baseFileName + ".xlsx"

	// File does not exist, create a new file
	f := excelize.NewFile()

	// Define sheet names
	sheetNames := []string{"MEMO", "BERITA ACARA", "SK", "SURAT", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE", "JADWAL CUTI", "JADWAL RAPAT", "BOOKING RAPAT", "TIMELINE DESKTOP", "TIMELINE PROJECT"}

	// Create all sheets
	for _, sheetName := range sheetNames {
		f.NewSheet(sheetName)
	}

	// Fetch data from database
	var memos []models.Memo
	var beritaAcaras []models.BeritaAcara
	var sks []models.Sk
	var surats []models.Surat
	var projects []models.Project
	var perdins []models.Perdin
	var suratMasuks []models.SuratMasuk
	var suratKeluars []models.SuratKeluar
	var meetingLists []models.MeetingSchedule
	var meetings []models.Meeting
	var arsips []models.Arsip
	var jadwalCutis []models.JadwalCuti
	var jadwalRapats []models.JadwalRapat
	var bookingRapats []models.BookingRapat
	var timelineDesktops []models.TimelineDesktop
	var timelineProjects []models.TimelineProject

	initializers.DB.Find(&memos)
	initializers.DB.Find(&beritaAcaras)
	initializers.DB.Find(&sks)
	initializers.DB.Find(&surats)
	initializers.DB.Find(&projects)
	initializers.DB.Find(&perdins)
	initializers.DB.Find(&suratMasuks)
	initializers.DB.Find(&suratKeluars)
	initializers.DB.Find(&meetingLists)
	initializers.DB.Find(&meetings)
	initializers.DB.Find(&arsips)
	initializers.DB.Find(&jadwalCutis)
	initializers.DB.Find(&jadwalRapats)
	initializers.DB.Find(&bookingRapats)
	initializers.DB.Find(&timelineDesktops)
	initializers.DB.Find(&timelineProjects)

	// Update data in each sheet
	for _, sheetName := range sheetNames {
		// Write header row
		switch sheetName {
		case "MEMO":

			// Header untuk SAG (kolom kiri)
			f.SetCellValue(sheetName, "A1", "Tanggal")
			f.SetCellValue(sheetName, "B1", "No Surat")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "PIC")

			// Header untuk ISO (kolom kanan)
			f.SetCellValue(sheetName, "F1", "Tanggal")
			f.SetCellValue(sheetName, "G1", "No Surat")
			f.SetCellValue(sheetName, "H1", "Perihal")
			f.SetCellValue(sheetName, "I1", "PIC")

		case "PROJECT":
			f.SetCellValue(sheetName, "A1", "Kode Project")
			f.SetCellValue(sheetName, "B1", "Jenis Pengadaan")
			f.SetCellValue(sheetName, "C1", "Nama Pengadaan")
			f.SetCellValue(sheetName, "D1", "Divisi Inisiasi")
			f.SetCellValue(sheetName, "E1", "Bulan")
			f.SetCellValue(sheetName, "F1", "Sumber Pendanaan")
			f.SetCellValue(sheetName, "G1", "Anggaran")
			f.SetCellValue(sheetName, "H1", "No Izin")
			f.SetCellValue(sheetName, "I1", "Tgl Izin")
			f.SetCellValue(sheetName, "J1", "Tgl TOR")
			f.SetCellValue(sheetName, "K1", "Pic")
			f.SetCellValue(sheetName, "F2", "SAG")

			// style Header Project
			styleHeader, err := f.NewStyle(&excelize.Style{
				Fill:      excelize.Fill{Type: "pattern", Color: []string{"6EB6F8"}, Pattern: 1},
				Font:      &excelize.Font{Bold: true, Color: "000000", VertAlign: "center"},
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
				},
			})
			if err != nil {
				return
			}
			err = f.SetCellStyle("PROJECT", fmt.Sprintf("A1"), fmt.Sprintf("K1"), styleHeader)

		case "BERITA ACARA":

			// Header untuk SAG (kolom kiri)
			f.SetCellValue(sheetName, "A1", "Tanggal")
			f.SetCellValue(sheetName, "B1", "No Surat")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "PIC")

			// Header untuk ISO (kolom kanan)
			f.SetCellValue(sheetName, "F1", "Tanggal")
			f.SetCellValue(sheetName, "G1", "No Surat")
			f.SetCellValue(sheetName, "H1", "Perihal")
			f.SetCellValue(sheetName, "I1", "PIC")

		case "SK":

			// Header untuk SAG (kolom kiri)
			f.SetCellValue(sheetName, "A1", "Tanggal")
			f.SetCellValue(sheetName, "B1", "No Surat")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "PIC")

			// Header untuk ISO (kolom kanan)
			f.SetCellValue(sheetName, "F1", "Tanggal")
			f.SetCellValue(sheetName, "G1", "No Surat")
			f.SetCellValue(sheetName, "H1", "Perihal")
			f.SetCellValue(sheetName, "I1", "PIC")

		case "SURAT":
			// Header untuk SAG (kolom kiri)
			f.SetCellValue(sheetName, "A1", "Tanggal")
			f.SetCellValue(sheetName, "B1", "No Surat")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "PIC")

			// Header untuk ISO (kolom kanan)
			f.SetCellValue(sheetName, "F1", "Tanggal")
			f.SetCellValue(sheetName, "G1", "No Surat")
			f.SetCellValue(sheetName, "H1", "Perihal")
			f.SetCellValue(sheetName, "I1", "PIC")

		case "PERDIN":

			f.SetCellValue(sheetName, "A1", "No Perdin")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Deskripsi")
			f.MergeCell(sheetName, "C1", "D1") // Menggabungkan sel C1 dan D1

			f.SetColWidth(sheetName, "A", "B", 20)
			f.SetColWidth(sheetName, "C", "D", 28)
			f.SetRowHeight(sheetName, 1, 28)

		case "SURAT MASUK":

			f.SetCellValue(sheetName, "A1", "No Surat")
			f.SetCellValue(sheetName, "B1", "Title Of Letter")
			f.SetCellValue(sheetName, "C1", "Related Divisi")
			f.SetCellValue(sheetName, "D1", "Destiny")
			f.SetCellValue(sheetName, "E1", "Date Issue")

			f.SetColWidth(sheetName, "A", "E", 20)

			styleHeader, err := f.NewStyle(&excelize.Style{
				Fill: excelize.Fill{
					Type:    "pattern",
					Color:   []string{"#4F81BD"},
					Pattern: 1,
				},
				Font: &excelize.Font{
					Bold:  true,
					Size:  12,
					Color: "FFFFFF",
				},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
				},
				Alignment: &excelize.Alignment{
					Horizontal: "center",
					Vertical:   "center",
				},
			})

			if err != nil {
				c.String(http.StatusInternalServerError, "Error creating style: %v", err)
				return
			}

			err = f.SetCellStyle("SURAT MASUK", "A1", "E1", styleHeader)

		case "SURAT KELUAR":
			f.SetCellValue(sheetName, "A1", "No Surat")
			f.SetCellValue(sheetName, "B1", "Title Of Letter")
			f.SetCellValue(sheetName, "C1", "From")
			f.SetCellValue(sheetName, "D1", "Pic")
			f.SetCellValue(sheetName, "E1", "Date Issue")

			styleHeader, err := f.NewStyle(&excelize.Style{
				Fill: excelize.Fill{
					Type:    "pattern",
					Color:   []string{"#4F81BD"},
					Pattern: 1,
				},
				Font: &excelize.Font{
					Bold:  true,
					Size:  12,
					Color: "FFFFFF",
				},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
				},
				Alignment: &excelize.Alignment{
					Horizontal: "center",
					Vertical:   "center",
				},
			})
			if err != nil {
				fmt.Println(err)
			}

			f.SetCellStyle(sheetName, "A1", "E1", styleHeader)

			f.SetColWidth(sheetName, "A", "A", 27)
			f.SetColWidth(sheetName, "B", "B", 40)
			f.SetColWidth(sheetName, "C", "C", 20)
			f.SetColWidth(sheetName, "D", "D", 20)
			f.SetColWidth(sheetName, "E", "E", 20)
			f.SetRowHeight(sheetName, 1, 20)
		case "MEETING":

			f.SetCellValue(sheetName, "A1", "TASK")
			f.SetCellValue(sheetName, "B1", "TINDAK LANJUT")
			f.SetCellValue(sheetName, "C1", "STATUS")
			f.SetCellValue(sheetName, "D1", "UPDATE PENGERJAAN")
			f.SetCellValue(sheetName, "E1", "PIC")
			f.SetCellValue(sheetName, "F1", "TANGGAL TARGET")
			f.SetCellValue(sheetName, "G1", "TANGGAL ACTUAL")

			f.SetColWidth("MEETING", "A", "A", 25)
			f.SetColWidth("MEETING", "B", "B", 40)
			f.SetColWidth("MEETING", "C", "C", 17)
			f.SetColWidth("MEETING", "D", "D", 27)
			f.SetColWidth("MEETING", "E", "E", 25)
			f.SetColWidth("MEETING", "F", "F", 20)
			f.SetColWidth("MEETING", "G", "G", 20)
			f.SetRowHeight("MEETING", 1, 35)

			FillColor, err := f.NewStyle(&excelize.Style{
				Fill: excelize.Fill{Type: "pattern", Color: []string{"eba55b"}, Pattern: 1},
				Alignment: &excelize.Alignment{
					Horizontal: "center",
					Vertical:   "center",
				},
				Border: []excelize.Border{
					{Type: "right", Color: "000000", Style: 1},
				},
			})
			if err != nil {
				fmt.Println(err)
			}
			err = f.SetCellStyle("MEETING", "A1", "G1", FillColor)

		case "MEETING SCHEDULE":
			f.SetCellValue(sheetName, "A1", "Hari")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "Waktu")
			f.SetCellValue(sheetName, "E1", "Selesai")
			f.SetCellValue(sheetName, "F1", "Tempat")
			f.SetCellValue(sheetName, "G1", "Status")
			f.SetCellValue(sheetName, "H1", "PIC")
			f.SetColWidth(sheetName, "A", "Z", 20)
		case "ARSIP":
			f.SetCellValue(sheetName, "A1", "No Arsip")
			f.SetCellValue(sheetName, "B1", "Jenis Dokumen")
			f.SetCellValue(sheetName, "C1", "No Dokumen")
			f.SetCellValue(sheetName, "D1", "Perihal")
			f.SetCellValue(sheetName, "E1", "No Box")
			f.SetCellValue(sheetName, "F1", "Keterangan")
			f.SetCellValue(sheetName, "G1", "Tanggal Dokumen")
			f.SetCellValue(sheetName, "H1", "Tanggal Penyerahan")
			f.SetColWidth(sheetName, "A", "Z", 20)
		case "JADWAL CUTI":
			// Ambil data dari model JadwalCuti
			var events []models.JadwalCuti
			if err := initializers.DB.Find(&events).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			sheet := "JADWAL CUTI"
			months := []string{
				"January 2024", "February 2024", "March 2024", "April 2024",
				"May 2024", "June 2024", "July 2024", "August 2024",
				"September 2024", "October 2024", "November 2024", "December 2024",
			}

			rowOffset := 0
			colOffset := 0
			for i, month := range months {
				setMonthDataCuti(f, sheet, month, rowOffset, colOffset, events)
				colOffset += 9 // Sesuaikan offset untuk bulan berikutnya dalam baris yang sama
				if (i+1)%3 == 0 {
					rowOffset += 18 // Pindah ke baris berikutnya setiap 3 bulan
					colOffset = 0
				}
			}
		case "JADWAL RAPAT":

			// Ambil data dari model JadwalRapat
			var events_rapat []models.JadwalRapat
			if err := initializers.DB.Find(&events_rapat).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			
			sheet := "JADWAL RAPAT"
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
		case "BOOKING RAPAT":
			// Ambil data dari model BookingRapat
			var events_rapat []models.BookingRapat
			if err := initializers.DB.Find(&events_rapat).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			
			sheet := "BOOKING RAPAT"
			months := []string{
				"January 2024", "February 2024", "March 2024", "April 2024",
				"May 2024", "June 2024", "July 2024", "August 2024",
				"September 2024", "October 2024", "November 2024", "December 2024",
			}

			rowOffset := 0
			colOffset := 0
			for i, month := range months {
				setMonthDataBookingRapat(f, sheet, month, rowOffset, colOffset, events_rapat)
				colOffset += 9 // Sesuaikan offset untuk bulan berikutnya dalam baris yang sama
				if (i+1)%3 == 0 {
					rowOffset += 18 // Pindah ke baris berikutnya setiap 3 bulan
					colOffset = 0
				}
			}
		case "TIMELINE DESKTOP":
			// Ambil data dari model TimelineDesktop
			var events_timeline []models.TimelineDesktop
			if err := initializers.DB.Find(&events_timeline).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			// Ambil data resources untuk digunakan dalam setMonthDataProject
			var resources []models.ResourceDesktop
			if err := initializers.DB.Find(&resources).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			resourceMap := make(map[uint]string)
			for _, resource := range resources {
				resourceMap[resource.ID] = resource.Name
			}

			
			sheet := "TIMELINE DESKTOP"
			months := []string{
				"January 2024", "February 2024", "March 2024", "April 2024",
				"May 2024", "June 2024", "July 2024", "August 2024",
				"September 2024", "October 2024", "November 2024", "December 2024",
			}

			rowOffset := 0
			colOffset := 0
			for i, month := range months {
				setMonthDataDesktop(f, sheet, month, rowOffset, colOffset, events_timeline, resourceMap)
				colOffset += 9 // Sesuaikan offset untuk bulan berikutnya dalam baris yang sama
				if (i+1)%3 == 0 {
					rowOffset += 18 // Pindah ke baris berikutnya setiap 3 bulan
					colOffset = 0
				}
			}

		case "TIMELINE PROJECT":
			// Ambil data dari model JadwalCuti
			var events_timeline []models.TimelineProject
			if err := initializers.DB.Find(&events_timeline).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			// Ambil data resources untuk digunakan dalam setMonthDataProject
			var resources []models.ResourceProject
			if err := initializers.DB.Find(&resources).Error; err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			resourceMap := make(map[uint]string)
			for _, resource := range resources {
				resourceMap[resource.ID] = resource.Name
			}

			
			sheet := "TIMELINE PROJECT"
			months := []string{
				"January 2024", "February 2024", "March 2024", "April 2024",
				"May 2024", "June 2024", "July 2024", "August 2024",
				"September 2024", "October 2024", "November 2024", "December 2024",
			}

			rowOffset := 0
			colOffset := 0
			for i, month := range months {
				setMonthDataProject(f, sheet, month, rowOffset, colOffset, events_timeline, resourceMap)
				colOffset += 9 // Sesuaikan offset untuk bulan berikutnya dalam baris yang sama
				if (i+1)%3 == 0 {
					rowOffset += 18 // Pindah ke baris berikutnya setiap 3 bulan
					colOffset = 0
				}
			}

		}

		// Write data rows
		var dataRows []interface{}
		switch sheetName {
		case "MEMO":
			for _, memo := range memos {
				dataRows = append(dataRows, memo)
			}
		case "BERITA ACARA":
			for _, beritaAcara := range beritaAcaras {
				dataRows = append(dataRows, beritaAcara)
			}
		case "SK":
			for _, sk := range sks {
				dataRows = append(dataRows, sk)
			}
		case "SURAT":
			for _, surat := range surats {
				dataRows = append(dataRows, surat)
			}
		case "PROJECT":
			for _, project := range projects {
				dataRows = append(dataRows, project)
			}
		case "PERDIN":
			for _, perdin := range perdins {
				dataRows = append(dataRows, perdin)
			}
		case "SURAT MASUK":
			for _, suratMasuk := range suratMasuks {
				dataRows = append(dataRows, suratMasuk)
			}
		case "SURAT KELUAR":
			for _, suratKeluar := range suratKeluars {
				dataRows = append(dataRows, suratKeluar)
			}
		case "MEETING":
			var meetings []models.Meeting
			initializers.DB.Find(&meetings)
			for _, meeting := range meetings {
				dataRows = append(dataRows, meeting)
			}
		case "MEETING SCHEDULE":
			var meetingLists []models.MeetingSchedule
			initializers.DB.Find(&meetingLists)
			for _, meetingList := range meetingLists {
				dataRows = append(dataRows, meetingList)
			}
		case "ARSIP":
			var arsips []models.Arsip
			initializers.DB.Find(&arsips)
			for _, arsip := range arsips {
				dataRows = append(dataRows, arsip)
			}
		}

		for i, dataRow := range dataRows {
			rowNum := i + 2 // Start from the second row (first row is header)
			switch sheetName {

			/******************************** Set Value Memo ***************************************/

			case "MEMO":
				// Inisialisasi baris awal
				rowSAG := 2
				rowISO := 2

				// Loop melalui data memo
				for _, memo := range memos {
					// Pastikan untuk dereferensikan pointer jika tidak nil
					var tanggal, noMemo, perihal, pic string
					if memo.Tanggal != nil {
						tanggal = memo.Tanggal.Format("2006-01-02") // Format tanggal sesuai kebutuhan
					}
					if memo.NoMemo != nil {
						noMemo = *memo.NoMemo
					}
					if memo.Perihal != nil {
						perihal = *memo.Perihal
					}
					if memo.Pic != nil {
						pic = *memo.Pic
					}

					// Pisahkan NoMemo untuk mendapatkan tipe memo
					parts := strings.Split(*memo.NoMemo, "/")
					if len(parts) > 1 && parts[1] == "ITS-SAG" {
						// Isi kolom SAG di sebelah kiri
						f.SetCellValue("MEMO", fmt.Sprintf("A%d", rowSAG), tanggal)
						f.SetCellValue("MEMO", fmt.Sprintf("B%d", rowSAG), noMemo)
						f.SetCellValue("MEMO", fmt.Sprintf("C%d", rowSAG), perihal)
						f.SetCellValue("MEMO", fmt.Sprintf("D%d", rowSAG), pic)
						rowSAG++
					} else if len(parts) > 1 && parts[1] == "ITS-ISO" {
						// Isi kolom ISO di sebelah kanan
						f.SetCellValue("MEMO", fmt.Sprintf("F%d", rowISO), tanggal)
						f.SetCellValue("MEMO", fmt.Sprintf("G%d", rowISO), noMemo)
						f.SetCellValue("MEMO", fmt.Sprintf("H%d", rowISO), perihal)
						f.SetCellValue("MEMO", fmt.Sprintf("I%d", rowISO), pic)
						rowISO++
					}
				}

				// style Line
				lastRowSAG := rowSAG - 1
				lastRowISO := rowISO - 1
				lastRow := lastRowSAG
				if lastRowISO > lastRowSAG {
					lastRow = lastRowISO
				}

				// Set lebar kolom agar rapi
				f.SetColWidth("MEMO", "A", "D", 20)
				f.SetColWidth("MEMO", "F", "I", 20)
				f.SetColWidth("MEMO", "E", "E", 2)
				for i := 2; i <= lastRow; i++ {
					f.SetRowHeight("MEMO", i, 30)
				}

				// style Line
				styleLine, err := f.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Border: []excelize.Border{
						{Type: "bottom", Color: "FFFFFF", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("MEMO", "E1", fmt.Sprintf("E%d", lastRow), styleLine)

				// style Border
				styleBorder, err := f.NewStyle(&excelize.Style{
					Border: []excelize.Border{
						{Type: "left", Color: "8E8E8E", Style: 2},
						{Type: "top", Color: "8E8E8E", Style: 2},
						{Type: "bottom", Color: "8E8E8E", Style: 2},
						{Type: "right", Color: "8E8E8E", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("MEMO", "A1", fmt.Sprintf("D%d", lastRow), styleBorder)
				err = f.SetCellStyle("MEMO", "F1", fmt.Sprintf("I%d", lastRow), styleBorder)

			/******************************** Set Value Berita Acara ***************************************/

			case "BERITA ACARA":
				// Inisialisasi baris awal
				rowSAG := 2
				rowISO := 2

				// Loop melalui data memo
				for _, beritaAcara := range beritaAcaras {
					// Pastikan untuk dereferensikan pointer jika tidak nil
					var tanggal, noSurat, perihal, pic string
					if beritaAcara.Tanggal != nil {
						tanggal = beritaAcara.Tanggal.Format("2006-01-02") // Format tanggal sesuai kebutuhan
					}
					if beritaAcara.NoSurat != nil {
						noSurat = *beritaAcara.NoSurat
					}
					if beritaAcara.Perihal != nil {
						perihal = *beritaAcara.Perihal
					}
					if beritaAcara.Pic != nil {
						pic = *beritaAcara.Pic
					}

					// Pisahkan NoMemo untuk mendapatkan tipe memo
					parts := strings.Split(*beritaAcara.NoSurat, "/")
					if len(parts) > 1 && parts[1] == "ITS-SAG" {
						// Isi kolom SAG di sebelah kiri
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("A%d", rowSAG), tanggal)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("B%d", rowSAG), noSurat)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("C%d", rowSAG), perihal)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("D%d", rowSAG), pic)
						rowSAG++
					} else if len(parts) > 1 && parts[1] == "ITS-ISO" {
						// Isi kolom ISO di sebelah kanan
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("F%d", rowISO), tanggal)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("G%d", rowISO), noSurat)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("H%d", rowISO), perihal)
						f.SetCellValue("BERITA ACARA", fmt.Sprintf("I%d", rowISO), pic)
						rowISO++
					}
				}

				// style Line
				lastRowSAG := rowSAG - 1
				lastRowISO := rowISO - 1
				lastRow := lastRowSAG
				if lastRowISO > lastRowSAG {
					lastRow = lastRowISO
				}

				// Set lebar kolom agar rapi
				f.SetColWidth("BERITA ACARA", "A", "D", 20)
				f.SetColWidth("BERITA ACARA", "F", "I", 20)
				f.SetColWidth("BERITA ACARA", "E", "E", 2)
				for i := 2; i <= lastRow; i++ {
					f.SetRowHeight("BERITA ACARA", i, 30)
				}

				// style Line
				styleLine, err := f.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Border: []excelize.Border{
						{Type: "bottom", Color: "FFFFFF", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("BERITA ACARA", "E1", fmt.Sprintf("E%d", lastRow), styleLine)

				// style Border
				styleBorder, err := f.NewStyle(&excelize.Style{
					Border: []excelize.Border{
						{Type: "left", Color: "8E8E8E", Style: 2},
						{Type: "top", Color: "8E8E8E", Style: 2},
						{Type: "bottom", Color: "8E8E8E", Style: 2},
						{Type: "right", Color: "8E8E8E", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("BERITA ACARA", "A1", fmt.Sprintf("D%d", lastRow), styleBorder)
				err = f.SetCellStyle("BERITA ACARA", "F1", fmt.Sprintf("I%d", lastRow), styleBorder)

			case "SK":
				// Inisialisasi baris awal
				rowSAG := 2
				rowISO := 2

				// Loop melalui data memo
				for _, sK := range sks {
					// Pastikan untuk dereferensikan pointer jika tidak nil
					var tanggal, noSurat, perihal, pic string
					if sK.Tanggal != nil {
						tanggal = sK.Tanggal.Format("2006-01-02") // Format tanggal sesuai kebutuhan
					}
					if sK.NoSurat != nil {
						noSurat = *sK.NoSurat
					}
					if sK.Perihal != nil {
						perihal = *sK.Perihal
					}
					if sK.Pic != nil {
						pic = *sK.Pic
					}

					// Pisahkan NoMemo untuk mendapatkan tipe memo
					parts := strings.Split(*sK.NoSurat, "/")
					if len(parts) > 1 && parts[1] == "ITS-SAG" {
						// Isi kolom SAG di sebelah kiri
						f.SetCellValue("SK", fmt.Sprintf("A%d", rowSAG), tanggal)
						f.SetCellValue("SK", fmt.Sprintf("B%d", rowSAG), noSurat)
						f.SetCellValue("SK", fmt.Sprintf("C%d", rowSAG), perihal)
						f.SetCellValue("SK", fmt.Sprintf("D%d", rowSAG), pic)
						rowSAG++
					} else if len(parts) > 1 && parts[1] == "ITS-ISO" {
						// Isi kolom ISO di sebelah kanan
						f.SetCellValue("SK", fmt.Sprintf("F%d", rowISO), tanggal)
						f.SetCellValue("SK", fmt.Sprintf("G%d", rowISO), noSurat)
						f.SetCellValue("SK", fmt.Sprintf("H%d", rowISO), perihal)
						f.SetCellValue("SK", fmt.Sprintf("I%d", rowISO), pic)
						rowISO++
					}
				}

				// style Line
				lastRowSAG := rowSAG - 1
				lastRowISO := rowISO - 1
				lastRow := lastRowSAG
				if lastRowISO > lastRowSAG {
					lastRow = lastRowISO
				}

				// Set lebar kolom agar rapi
				f.SetColWidth("SK", "A", "D", 20)
				f.SetColWidth("SK", "F", "I", 20)
				f.SetColWidth("SK", "E", "E", 2)
				for i := 2; i <= lastRow; i++ {
					f.SetRowHeight("SK", i, 30)
				}

				// style Line
				styleLine, err := f.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Border: []excelize.Border{
						{Type: "bottom", Color: "FFFFFF", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("SK", "E1", fmt.Sprintf("E%d", lastRow), styleLine)

				// style Border
				styleBorder, err := f.NewStyle(&excelize.Style{
					Border: []excelize.Border{
						{Type: "left", Color: "8E8E8E", Style: 2},
						{Type: "top", Color: "8E8E8E", Style: 2},
						{Type: "bottom", Color: "8E8E8E", Style: 2},
						{Type: "right", Color: "8E8E8E", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("SK", "A1", fmt.Sprintf("D%d", lastRow), styleBorder)
				err = f.SetCellStyle("SK", "F1", fmt.Sprintf("I%d", lastRow), styleBorder)

			case "SURAT":
				// Inisialisasi baris awal
				rowSAG := 2
				rowISO := 2

				// Loop melalui data memo
				for _, surat := range surats {
					// Pastikan untuk dereferensikan pointer jika tidak nil
					var tanggal, noSurat, perihal, pic string
					if surat.Tanggal != nil {
						tanggal = surat.Tanggal.Format("2006-01-02") // Format tanggal sesuai kebutuhan
					}
					if surat.NoSurat != nil {
						noSurat = *surat.NoSurat
					}
					if surat.Perihal != nil {
						perihal = *surat.Perihal
					}
					if surat.Pic != nil {
						pic = *surat.Pic
					}

					// Pisahkan NoMemo untuk mendapatkan tipe memo
					parts := strings.Split(*surat.NoSurat, "/")
					if len(parts) > 1 && parts[1] == "ITS-SAG" {
						// Isi kolom SAG di sebelah kiri
						f.SetCellValue("SURAT", fmt.Sprintf("A%d", rowSAG), tanggal)
						f.SetCellValue("SURAT", fmt.Sprintf("B%d", rowSAG), noSurat)
						f.SetCellValue("SURAT", fmt.Sprintf("C%d", rowSAG), perihal)
						f.SetCellValue("SURAT", fmt.Sprintf("D%d", rowSAG), pic)
						rowSAG++
					} else if len(parts) > 1 && parts[1] == "ITS-ISO" {
						// Isi kolom ISO di sebelah kanan
						f.SetCellValue("SURAT", fmt.Sprintf("F%d", rowISO), tanggal)
						f.SetCellValue("SURAT", fmt.Sprintf("G%d", rowISO), noSurat)
						f.SetCellValue("SURAT", fmt.Sprintf("H%d", rowISO), perihal)
						f.SetCellValue("SURAT", fmt.Sprintf("I%d", rowISO), pic)
						rowISO++
					}
				}

				// style Line
				lastRowSAG := rowSAG - 1
				lastRowISO := rowISO - 1
				lastRow := lastRowSAG
				if lastRowISO > lastRowSAG {
					lastRow = lastRowISO
				}

				// Set lebar kolom agar rapi
				f.SetColWidth("SURAT", "A", "D", 20)
				f.SetColWidth("SURAT", "F", "I", 20)
				f.SetColWidth("SURAT", "E", "E", 2)
				for i := 2; i <= lastRow; i++ {
					f.SetRowHeight("SURAT", i, 30)
				}

				// style Line
				styleLine, err := f.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Border: []excelize.Border{
						{Type: "bottom", Color: "FFFFFF", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("SURAT", "E1", fmt.Sprintf("E%d", lastRow), styleLine)

				// style Border
				styleBorder, err := f.NewStyle(&excelize.Style{
					Border: []excelize.Border{
						{Type: "left", Color: "8E8E8E", Style: 2},
						{Type: "top", Color: "8E8E8E", Style: 2},
						{Type: "bottom", Color: "8E8E8E", Style: 2},
						{Type: "right", Color: "8E8E8E", Style: 2},
					},
				})
				if err != nil {
					fmt.Println(err)
				}
				err = f.SetCellStyle("SURAT", "A1", fmt.Sprintf("D%d", lastRow), styleBorder)
				err = f.SetCellStyle("SURAT", "F1", fmt.Sprintf("I%d", lastRow), styleBorder)

			case "PROJECT":

				// Set style for column B
				styleLine, err := f.NewStyle(&excelize.Style{
					Fill:      excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Font:      &excelize.Font{Bold: true, Color: "FFFFFF", VertAlign: "center"},
					Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
					},
				})
				if err != nil {
					return
				}
				err = f.SetCellStyle("PROJECT", fmt.Sprintf("A2"), fmt.Sprintf("K2"), styleLine)
				if err != nil {
					return
				}

				// Initialize row counters
				rowIndex := 3
				lastRowSAG := 3

				styleAll, err := f.NewStyle(&excelize.Style{
					Alignment: &excelize.Alignment{WrapText: true},
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
					},
				})
				if err != nil {
					return
				}

				// Loop through projects
				for _, project := range projects {
					// Dereference pointers if not nil
					var kodeProject, jenisPengadaan, namaPengadaan, divInisiasi, bulan, sumberPendanaan, noIzin, tanggalIzin, tanggalTor, pic string
					if project.KodeProject != nil {
						kodeProject = *project.KodeProject
					}
					if project.JenisPengadaan != nil {
						jenisPengadaan = *project.JenisPengadaan
					}
					if project.NamaPengadaan != nil {
						namaPengadaan = *project.NamaPengadaan
					}
					if project.DivInisiasi != nil {
						divInisiasi = *project.DivInisiasi
					}
					if project.Bulan != nil {
						bulan = project.Bulan.Format("Jan-06")
					}
					if project.SumberPendanaan != nil {
						sumberPendanaan = *project.SumberPendanaan
					}
					if project.NoIzin != nil {
						noIzin = *project.NoIzin
					}
					if project.TanggalIzin != nil {
						tanggalIzin = project.TanggalIzin.Format("2006-01-02")
					}
					if project.TanggalTor != nil {
						tanggalTor = project.TanggalTor.Format("2006-01-02")
					}
					if project.Pic != nil {
						pic = *project.Pic
					}

					// Split NoMemo to get memo type
					parts := strings.Split(*project.KodeProject, "/")
					if len(parts) > 1 && parts[1] == "ITS-SAG" {
						// Fill SAG column (left)
						f.SetCellValue("PROJECT", fmt.Sprintf("A%d", rowIndex), kodeProject)
						f.SetCellValue("PROJECT", fmt.Sprintf("B%d", rowIndex), jenisPengadaan)
						f.SetCellValue("PROJECT", fmt.Sprintf("C%d", rowIndex), namaPengadaan)
						f.SetCellValue("PROJECT", fmt.Sprintf("D%d", rowIndex), divInisiasi)
						f.SetCellValue("PROJECT", fmt.Sprintf("E%d", rowIndex), bulan)
						f.SetCellValue("PROJECT", fmt.Sprintf("F%d", rowIndex), sumberPendanaan)

						if project.Anggaran != nil {
							anggaran, err := strconv.ParseInt(*project.Anggaran, 10, 64)
							if err != nil {
								return
							}
							formattedAnggaran := fmt.Sprintf("Rp. %d", anggaran)
							f.SetCellValue("PROJECT", fmt.Sprintf("G%d", rowIndex), formattedAnggaran)
							styleAnggaran, err := f.NewStyle(&excelize.Style{
								NumFmt: 3,
							})
							if err != nil {
								return
							}
							err = f.SetCellStyle("PROJECT", fmt.Sprintf("G%d", rowIndex), fmt.Sprintf("G%d", rowIndex), styleAnggaran)
							if err != nil {
								return
							}
						}

						f.SetCellValue("PROJECT", fmt.Sprintf("H%d", rowIndex), noIzin)
						f.SetCellValue("PROJECT", fmt.Sprintf("I%d", rowIndex), tanggalIzin)
						f.SetCellValue("PROJECT", fmt.Sprintf("J%d", rowIndex), tanggalTor)
						f.SetCellValue("PROJECT", fmt.Sprintf("K%d", rowIndex), pic)
						err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", rowIndex), fmt.Sprintf("K%d", rowIndex), styleAll)
						if err != nil {
							return
						}
						rowIndex++
						lastRowSAG = rowIndex
					}

					if len(parts) > 1 && parts[1] == "ITS-ISO" {
						rowISO := rowIndex + 1
						// Fill ISO column (right)
						f.SetCellValue("PROJECT", fmt.Sprintf("A%d", rowISO), kodeProject)
						f.SetCellValue("PROJECT", fmt.Sprintf("B%d", rowISO), jenisPengadaan)
						f.SetCellValue("PROJECT", fmt.Sprintf("C%d", rowISO), namaPengadaan)
						f.SetCellValue("PROJECT", fmt.Sprintf("D%d", rowISO), divInisiasi)
						f.SetCellValue("PROJECT", fmt.Sprintf("E%d", rowISO), bulan)
						f.SetCellValue("PROJECT", fmt.Sprintf("F%d", rowISO), sumberPendanaan)

						if project.Anggaran != nil {
							anggaran, err := strconv.ParseInt(*project.Anggaran, 10, 64)
							if err != nil {
								return
							}
							formattedAnggaran := fmt.Sprintf("Rp. %d", anggaran)
							f.SetCellValue("PROJECT", fmt.Sprintf("G%d", rowISO), formattedAnggaran)
							styleAnggaran, err := f.NewStyle(&excelize.Style{
								NumFmt: 3,
							})
							if err != nil {
								return
							}
							err = f.SetCellStyle("PROJECT", fmt.Sprintf("G%d", rowISO), fmt.Sprintf("G%d", rowISO), styleAnggaran)
							if err != nil {
								return
							}
						}

						f.SetCellValue("PROJECT", fmt.Sprintf("H%d", rowISO), noIzin)
						f.SetCellValue("PROJECT", fmt.Sprintf("I%d", rowISO), tanggalIzin)
						f.SetCellValue("PROJECT", fmt.Sprintf("J%d", rowISO), tanggalTor)
						f.SetCellValue("PROJECT", fmt.Sprintf("K%d", rowISO), pic)
						err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", rowISO), fmt.Sprintf("K%d", rowISO), styleAll)
						if err != nil {
							return
						}
						rowIndex++
					}
				}

				for i := 3; i < lastRowSAG; i++ {
					f.SetRowHeight("PROJECT", i, 30)
				}

				// Set black line after SAG data is generated
				styleLine, err = f.NewStyle(&excelize.Style{
					Fill:      excelize.Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
					Font:      &excelize.Font{Bold: true, Color: "FFFFFF", VertAlign: "center"},
					Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
					},
				})
				if err != nil {
					return
				}

				f.SetCellValue("PROJECT", fmt.Sprintf("F%d", lastRowSAG), "ISO")
				err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", lastRowSAG), fmt.Sprintf("K%d", lastRowSAG), styleLine)
				if err != nil {
					return
				}

				// Set column widths after all data is filled
				f.SetColWidth("PROJECT", "A", "A", 30)
				f.SetColWidth("PROJECT", "B", "B", 15)
				f.SetColWidth("PROJECT", "C", "C", 35)
				f.SetColWidth("PROJECT", "D", "D", 22)
				f.SetColWidth("PROJECT", "E", "E", 15)
				f.SetColWidth("PROJECT", "F", "F", 20)
				f.SetColWidth("PROJECT", "G", "G", 20)
				f.SetColWidth("PROJECT", "H", "H", 23)
				f.SetColWidth("PROJECT", "I", "I", 22)
				f.SetColWidth("PROJECT", "J", "J", 20)
				f.SetColWidth("PROJECT", "K", "K", 16)

			case "PERDIN":
				// Fetch initial data from the database
				var perdins []models.Perdin
				initializers.DB.Find(&perdins)

				// Write initial data to the "PERDIN" sheet
				perdinSheetName := "PERDIN"
				for i, perdin := range perdins {
					var tanggalString string
					if perdin.Tanggal == nil {
						tanggalString = "" // Atau nilai default lain yang Anda inginkan
					} else {
						tanggalString = perdin.Tanggal.Format("2006-01-02")
					}
					rowNum := i + 2 // Start from the second row (first row is header)

					f.SetCellValue(perdinSheetName, fmt.Sprintf("A%d", rowNum), getStringValue(perdin.NoPerdin))
					f.SetCellValue(perdinSheetName, fmt.Sprintf("B%d", rowNum), tanggalString)
					f.SetCellValue(perdinSheetName, fmt.Sprintf("C%d", rowNum), getStringValue(perdin.Hotel))
					f.SetCellValue(perdinSheetName, fmt.Sprintf("D%d", rowNum), getStringValue(perdin.Transport))

					f.SetRowHeight(perdinSheetName, rowNum, 15)
				}

				// Apply border to all cells
				style, err := f.NewStyle(&excelize.Style{
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
					},
				})
				if err != nil {
					c.String(http.StatusInternalServerError, "Error creating style: %v", err)
					return
				}

				// Apply the style to the entire sheet
				f.SetCellStyle(perdinSheetName, "A1", fmt.Sprintf("D%d", len(perdins)+1), style)

				styleHeader, err := f.NewStyle(&excelize.Style{
					Alignment: &excelize.Alignment{
						Horizontal: "center",
						Vertical:   "center",
					},
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
					},
				})
				if err != nil {
					c.String(http.StatusInternalServerError, "Error creating style: %v", err)
					return
				}

				// Apply the style to the entire sheet
				f.SetCellStyle(perdinSheetName, "A1", "D1", styleHeader)

			case "SURAT MASUK":

				// Fetch initial data from the database
				var surat_masuks []models.SuratMasuk
				initializers.DB.Find(&surat_masuks)

				// Write initial data to the "SURAT MASUK" sheet
				surat_masukSheetName := "SURAT MASUK"
				for i, surat_masuk := range surat_masuks {
					tanggalString := surat_masuk.Tanggal.Format("2 January 2006")
					rowNum := i + 2 // Start from the second row (first row is header)
					f.SetCellValue(surat_masukSheetName, fmt.Sprintf("A%d", rowNum), *surat_masuk.NoSurat)
					f.SetCellValue(surat_masukSheetName, fmt.Sprintf("B%d", rowNum), *surat_masuk.Title)
					f.SetCellValue(surat_masukSheetName, fmt.Sprintf("C%d", rowNum), *surat_masuk.RelatedDiv)
					f.SetCellValue(surat_masukSheetName, fmt.Sprintf("D%d", rowNum), *surat_masuk.DestinyDiv)
					f.SetCellValue(surat_masukSheetName, fmt.Sprintf("E%d", rowNum), tanggalString)

					f.SetColWidth("SURAT MASUK", "A", "A", 27)
					f.SetColWidth("SURAT MASUK", "B", "B", 40)
					f.SetColWidth("SURAT MASUK", "C", "C", 20)
					f.SetColWidth("SURAT MASUK", "D", "D", 20)
					f.SetColWidth("SURAT MASUK", "E", "E", 20)
					f.SetRowHeight("SURAT MASUK", 1, 20)

					styleData, err := f.NewStyle(&excelize.Style{
						Border: []excelize.Border{
							{Type: "left", Color: "000000", Style: 1},
							{Type: "top", Color: "000000", Style: 1},
							{Type: "bottom", Color: "000000", Style: 1},
							{Type: "right", Color: "000000", Style: 1},
						},
					})
					if err != nil {
						c.String(http.StatusInternalServerError, "Error creating style: %v", err)
						return
					}
					err = f.SetCellStyle(surat_masukSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("E%d", rowNum), styleData)
				}

			case "SURAT KELUAR":

				// Fetch initial data from the database
				var surat_keluars []models.SuratKeluar
				initializers.DB.Find(&surat_keluars)

				// Write initial data to the "SAG" sheet
				surat_keluarSheetName := "SURAT KELUAR"
				for i, surat_keluar := range surat_keluars {
					tanggalString := surat_keluar.Tanggal.Format("2 January 2006")
					rowNum := i + 2 // Start from the second row (first row is header)
					f.SetCellValue(surat_keluarSheetName, fmt.Sprintf("A%d", rowNum), *surat_keluar.NoSurat)
					f.SetCellValue(surat_keluarSheetName, fmt.Sprintf("B%d", rowNum), *surat_keluar.Title)
					f.SetCellValue(surat_keluarSheetName, fmt.Sprintf("C%d", rowNum), *surat_keluar.From)
					f.SetCellValue(surat_keluarSheetName, fmt.Sprintf("D%d", rowNum), *surat_keluar.Pic)
					f.SetCellValue(surat_keluarSheetName, fmt.Sprintf("E%d", rowNum), tanggalString)

					styleData, err := f.NewStyle(&excelize.Style{
						Border: []excelize.Border{
							{Type: "left", Color: "000000", Style: 1},
							{Type: "top", Color: "000000", Style: 1},
							{Type: "bottom", Color: "000000", Style: 1},
							{Type: "right", Color: "000000", Style: 1},
						},
					})
					if err != nil {
						fmt.Println(err)
					}

					f.SetCellStyle(surat_keluarSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("E%d", rowNum), styleData)
				}

			case "MEETING":

				wrapstyle, err := f.NewStyle(&excelize.Style{
					Alignment: &excelize.Alignment{
						WrapText: true,
						Vertical: "center",
					},
				})
				if err != nil {
					fmt.Println(err)
				}

				// Fetch initial data from the database
				var meetings []models.Meeting
				initializers.DB.Find(&meetings)

				err = f.SetCellStyle("MEETING", "A2", fmt.Sprintf("G%d", len(meetings)+1), wrapstyle)

				// Write initial data to the "MEETING" sheet
				meetingSheetName := "MEETING"
				for i, meeting := range meetings {
					tanggalTargetString := meeting.TanggalTarget.Format("2006-01-02")
					tanggalActualString := meeting.TanggalActual.Format("2006-01-02")
					rowNum := i + 2 // Start from the second row (first row is header)

					// Check for nil pointers and use the actual values
					task := ""
					if meeting.Task != nil {
						task = *meeting.Task
					}
					tindakLanjut := ""
					if meeting.TindakLanjut != nil {
						tindakLanjut = *meeting.TindakLanjut
					}
					status := ""
					if meeting.Status != nil {
						status = *meeting.Status
					}
					updatePengerjaan := ""
					if meeting.UpdatePengerjaan != nil {
						updatePengerjaan = *meeting.UpdatePengerjaan
					}
					pic := ""
					if meeting.Pic != nil {
						pic = *meeting.Pic
					}

					f.SetCellValue(meetingSheetName, fmt.Sprintf("A%d", rowNum), task)
					f.SetCellValue(meetingSheetName, fmt.Sprintf("B%d", rowNum), tindakLanjut)
					f.SetCellValue(meetingSheetName, fmt.Sprintf("C%d", rowNum), status) // Set status value

					// Apply styles based on status
					var styleID int
					switch status {
					case "Done":
						styleID, err = f.NewStyle(&excelize.Style{
							Font: &excelize.Font{
								Color: "00ff80",
							},
							Alignment: &excelize.Alignment{
								Horizontal: "center",
								Vertical:   "center",
							},
							Border: []excelize.Border{
								{Type: "left", Color: "000000", Style: 1},
								{Type: "top", Color: "000000", Style: 1},
								{Type: "bottom", Color: "000000", Style: 1},
								{Type: "right", Color: "000000", Style: 1},
							},
						})
					case "On Progress":
						styleID, err = f.NewStyle(&excelize.Style{
							Font: &excelize.Font{
								Color: "ffa500",
							},
							Alignment: &excelize.Alignment{
								Horizontal: "center",
								Vertical:   "center",
							},
							Border: []excelize.Border{
								{Type: "left", Color: "000000", Style: 1},
								{Type: "top", Color: "000000", Style: 1},
								{Type: "bottom", Color: "000000", Style: 1},
								{Type: "right", Color: "000000", Style: 1},
							},
						})
					case "Cancel":
						styleID, err = f.NewStyle(&excelize.Style{
							Font: &excelize.Font{
								Color: "FF3131",
							},
							Alignment: &excelize.Alignment{
								Horizontal: "center",
								Vertical:   "center",
							},
							Border: []excelize.Border{
								{Type: "left", Color: "000000", Style: 1},
								{Type: "top", Color: "000000", Style: 1},
								{Type: "bottom", Color: "000000", Style: 1},
								{Type: "right", Color: "000000", Style: 1},
							},
						})
					default:
						styleID, err = f.NewStyle(&excelize.Style{
							Border: []excelize.Border{
								{Type: "left", Color: "000000", Style: 1},
								{Type: "top", Color: "000000", Style: 1},
								{Type: "bottom", Color: "000000", Style: 1},
								{Type: "right", Color: "000000", Style: 1},
							},
						})
					}
					if err != nil {
						fmt.Println(err)
					}
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("C%d", rowNum), fmt.Sprintf("C%d", rowNum), styleID)

					// Apply border style to other cells
					borderStyle, err := f.NewStyle(&excelize.Style{
						Border: []excelize.Border{
							{Type: "left", Color: "000000", Style: 1},
							{Type: "top", Color: "000000", Style: 1},
							{Type: "bottom", Color: "000000", Style: 1},
							{Type: "right", Color: "000000", Style: 1},
						},
						Alignment: &excelize.Alignment{
							WrapText: true,
						},
					})
					if err != nil {
						fmt.Println(err)
					}
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("A%d", rowNum), borderStyle)
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("B%d", rowNum), fmt.Sprintf("B%d", rowNum), borderStyle)
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("D%d", rowNum), fmt.Sprintf("D%d", rowNum), borderStyle)
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("E%d", rowNum), fmt.Sprintf("E%d", rowNum), borderStyle)
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("F%d", rowNum), fmt.Sprintf("F%d", rowNum), borderStyle)
					f.SetCellStyle(meetingSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), borderStyle)

					f.SetCellValue(meetingSheetName, fmt.Sprintf("D%d", rowNum), updatePengerjaan)
					f.SetCellValue(meetingSheetName, fmt.Sprintf("E%d", rowNum), pic)
					f.SetCellValue(meetingSheetName, fmt.Sprintf("F%d", rowNum), tanggalTargetString)
					f.SetCellValue(meetingSheetName, fmt.Sprintf("G%d", rowNum), tanggalActualString)

					// Calculate row height based on content length
					maxContentLength := max(len(task), len(tindakLanjut), len(status), len(updatePengerjaan), len(pic))
					rowHeight := calculateRowHeight(maxContentLength)
					f.SetRowHeight(meetingSheetName, rowNum, rowHeight)
				}

			case "MEETING SCHEDULE":
				meetingList := dataRow.(models.MeetingSchedule)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), meetingList.Hari)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), meetingList.Tanggal.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), meetingList.Perihal)

				// Handle Waktu
				if meetingList.Waktu != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), *meetingList.Waktu)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), "")
				}

				// Handle Selesai
				if meetingList.Selesai != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), *meetingList.Selesai)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), "")
				}

				if meetingList.Tempat != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), *meetingList.Tempat)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), "")
				}

				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), meetingList.Status)
				f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), meetingList.Pic)
			case "ARSIP":
				arsip := dataRow.(models.Arsip)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), arsip.NoArsip)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), arsip.JenisDokumen)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), arsip.NoDokumen)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), arsip.Perihal)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), arsip.NoBox)
				f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), arsip.Keterangan)
				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), arsip.TanggalDokumen.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), arsip.TanggalPenyerahan.Format("2006-01-02"))
			}
		}
	}

	// Delete the default "Sheet1" sheet
	if err := f.DeleteSheet("Sheet1"); err != nil {
		fmt.Println("Error deleting Sheet1:", err) // Tambahkan log error
		// Handle error jika bukan error "sheet tidak ditemukan"
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		c.String(http.StatusInternalServerError, "Error saving file: %v", err)
		return
	}

	// Serve the file to the client
	c.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=%s", fileName))
	c.Writer.Write(buf.Bytes())
}

/* INI FUNCTION COBA COCA */

/******************************** INI FUNCTION UPDATE ALL SHEET ***************************************/

func UpdateAllSheets(c *gin.Context) {
	dir := "C:\\excel"
	fileName := "its_report.xlsx"
	filePath := filepath.Join(dir, fileName)

	// Check if the file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		c.String(http.StatusBadRequest, "File tidak ada")
		return
	}

	// Open the existing Excel file
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error membuka file: %v", err)
		return
	}
	defer f.Close()

	// Define sheet names
	sheetNames := []string{"MEMO", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}

	// Fetch updated data from the database
	var memos []models.Memo
	var projects []models.Project
	var perdins []models.Perdin
	var suratMasuks []models.SuratMasuk
	var suratKeluars []models.SuratKeluar
	var meetingLists []models.MeetingSchedule
	var meetings []models.Meeting
	var arsips []models.Arsip

	initializers.DB.Find(&memos)
	initializers.DB.Find(&projects)
	initializers.DB.Find(&perdins)
	initializers.DB.Find(&suratMasuks)
	initializers.DB.Find(&suratKeluars)
	initializers.DB.Find(&meetingLists)
	initializers.DB.Find(&meetings)
	initializers.DB.Find(&arsips)

	for _, sheetName := range sheetNames {
		// Write header row
		switch sheetName {
		case "MEMO":
			f.SetCellValue(sheetName, "A1", "Tanggal")
			f.SetCellValue(sheetName, "B1", "No Memo")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "Kategori")
			f.SetCellValue(sheetName, "E1", "PIC")

			f.SetColWidth(sheetName, "A", "E", 20)
		case "PROJECT":
			f.SetCellValue(sheetName, "A1", "Kode Project")
			f.SetCellValue(sheetName, "B1", "Jenis Pengadaan")
			f.SetCellValue(sheetName, "C1", "Nama Pengadaan")
			f.SetCellValue(sheetName, "D1", "Divisi Inisiasi")
			f.SetCellValue(sheetName, "E1", "Bulan")
			f.SetCellValue(sheetName, "F1", "Sumber Pendanaan")
			f.SetCellValue(sheetName, "G1", "Anggaran")
			f.SetCellValue(sheetName, "H1", "No Izin")
			f.SetCellValue(sheetName, "I1", "Tgl Izin")
			f.SetCellValue(sheetName, "J1", "Tgl TOR")
			f.SetCellValue(sheetName, "K1", "Pic")

			f.SetColWidth(sheetName, "A", "K", 20)
		case "PERDIN":
			f.SetCellValue(sheetName, "A1", "No Perdin")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Deskripsi")
			f.MergeCell(sheetName, "C1", "D1")

			f.SetColWidth(sheetName, "A", "D", 20)
		case "SURAT MASUK":
			f.SetCellValue(sheetName, "A1", "No Surat")
			f.SetCellValue(sheetName, "B1", "Title")
			f.SetCellValue(sheetName, "C1", "Related Divisi")
			f.SetCellValue(sheetName, "D1", "Destiny Divisi")
			f.SetCellValue(sheetName, "E1", "Tanggal")

			f.SetColWidth(sheetName, "A", "E", 20)
		case "SURAT KELUAR":
			f.SetCellValue(sheetName, "A1", "No Surat")
			f.SetCellValue(sheetName, "B1", "Title")
			f.SetCellValue(sheetName, "C1", "From")
			f.SetCellValue(sheetName, "D1", "Pic")
			f.SetCellValue(sheetName, "E1", "Tanggal")

			f.SetColWidth(sheetName, "A", "E", 20)
		case "ARSIP":
			f.SetCellValue(sheetName, "A1", "No Arsip")
			f.SetCellValue(sheetName, "B1", "Jenis Dokumen")
			f.SetCellValue(sheetName, "C1", "No Dokumen")
			f.SetCellValue(sheetName, "D1", "Perihal")
			f.SetCellValue(sheetName, "E1", "No Box")
			f.SetCellValue(sheetName, "F1", "Keterangan")
			f.SetCellValue(sheetName, "G1", "Tanggal Dokumen")
			f.SetCellValue(sheetName, "H1", "Tanggal Penyerahan")

			f.SetColWidth(sheetName, "A", "H", 20)
		case "MEETING":
			f.SetCellValue(sheetName, "A1", "TASK")
			f.SetCellValue(sheetName, "B1", "TINDAK LANJUT")
			f.SetCellValue(sheetName, "C1", "STATUS")
			f.SetCellValue(sheetName, "D1", "UPDATE PENGERJAAN")
			f.SetCellValue(sheetName, "E1", "PIC")
			f.SetCellValue(sheetName, "F1", "TANGGAL TARGET")
			f.SetCellValue(sheetName, "G1", "TANGGAL ACTUAL")

			f.SetColWidth(sheetName, "A", "G", 20)
		case "MEETING SCHEDULE":
			f.SetCellValue(sheetName, "A1", "Hari")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "Waktu")
			f.SetCellValue(sheetName, "E1", "Selesai")
			f.SetCellValue(sheetName, "F1", "Tempat")
			f.SetCellValue(sheetName, "G1", "Status")
			f.SetCellValue(sheetName, "H1", "PIC")

			f.SetColWidth(sheetName, "A", "H", 20)
		}

		// Write data rows
		var dataRows []interface{}
		switch sheetName {
		case "MEMO":
			for _, memo := range memos {
				dataRows = append(dataRows, memo)
			}
		case "PROJECT":
			for _, project := range projects {
				dataRows = append(dataRows, project)
			}
		case "PERDIN":
			for _, perdin := range perdins {
				dataRows = append(dataRows, perdin)
			}
		case "SURAT MASUK":
			for _, suratMasuk := range suratMasuks {
				dataRows = append(dataRows, suratMasuk)
			}
		case "SURAT KELUAR":
			for _, suratKeluar := range suratKeluars {
				dataRows = append(dataRows, suratKeluar)
			}
		case "ARSIP":
			for _, arsip := range arsips {
				dataRows = append(dataRows, arsip)
			}
		case "MEETING":
			for _, meeting := range meetings {
				dataRows = append(dataRows, meeting)
			}
		case "MEETING SCHEDULE":
			for _, meetingList := range meetingLists {
				dataRows = append(dataRows, meetingList)
			}

		}

		for i, dataRow := range dataRows {
			rowNum := i + 2 // Start from the second row (first row is header)
			switch sheetName {
			case "MEMO":
				memo := dataRow.(models.Memo)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), memo.Tanggal.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), memo.NoMemo)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), memo.Perihal)
				// f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), memo.Kategori)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), memo.Pic)

				f.SetColWidth(sheetName, "A", "E", 20)
			case "PROJECT":
				project := dataRow.(models.Project)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), project.KodeProject)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), project.JenisPengadaan)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), project.NamaPengadaan)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), project.DivInisiasi)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), project.Bulan.Format("02-01-2006")) // Write month as text
				f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), project.SumberPendanaan)
				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), project.Anggaran)
				f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), project.NoIzin)
				f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowNum), project.TanggalIzin.Format("02-01-2006"))
				f.SetCellValue(sheetName, fmt.Sprintf("J%d", rowNum), project.TanggalTor.Format("02-01-2006"))
				f.SetCellValue(sheetName, fmt.Sprintf("K%d", rowNum), project.Pic)

				f.SetColWidth(sheetName, "A", "K", 20)
			case "PERDIN":
				perdin := dataRow.(models.Perdin)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), perdin.NoPerdin)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), perdin.Tanggal.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), perdin.Hotel)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), perdin.Transport)

				f.SetColWidth(sheetName, "A", "D", 20)
			case "SURAT MASUK":
				suratMasuk := dataRow.(models.SuratMasuk)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), suratMasuk.NoSurat)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), suratMasuk.Title)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), suratMasuk.RelatedDiv)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), suratMasuk.DestinyDiv)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), suratMasuk.Tanggal.Format("2006-01-02"))

				f.SetColWidth(sheetName, "A", "E", 20)
			case "SURAT KELUAR":
				suratKeluar := dataRow.(models.SuratKeluar)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), suratKeluar.NoSurat)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), suratKeluar.Title)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), suratKeluar.From)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), suratKeluar.Pic)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), suratKeluar.Tanggal.Format("2006-01-02"))

				f.SetColWidth(sheetName, "A", "E", 20)
			case "ARSIP":
				arsip := dataRow.(models.Arsip)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), arsip.NoArsip)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), arsip.JenisDokumen)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), arsip.NoDokumen)
				f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), arsip.Perihal)
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), arsip.NoBox)
				f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), arsip.Keterangan)
				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), arsip.TanggalDokumen.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), arsip.TanggalPenyerahan.Format("2006-01-02"))

				f.SetColWidth(sheetName, "A", "H", 20)
			case "MEETING":
				meeting := dataRow.(models.Meeting)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), meeting.Task)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), meeting.TindakLanjut)
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), meeting.Status)
				if meeting.UpdatePengerjaan != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), *meeting.UpdatePengerjaan)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), "")
				}
				f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), meeting.Pic)
				f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), meeting.TanggalTarget.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), meeting.TanggalActual.Format("2006-01-02"))

				f.SetColWidth(sheetName, "A", "G", 20)
			case "MEETING SCHEDULE":
				meetingList := dataRow.(models.MeetingSchedule)
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), meetingList.Hari)
				f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), meetingList.Tanggal.Format("2006-01-02"))
				f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), meetingList.Perihal)

				// Handle Waktu
				if meetingList.Waktu != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), *meetingList.Waktu)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), "")
				}

				// Handle Selesai
				if meetingList.Selesai != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), *meetingList.Selesai)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), "")
				}

				if meetingList.Tempat != nil {
					f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), *meetingList.Tempat)
				} else {
					f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), "")
				}

				f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), meetingList.Status)
				f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), meetingList.Pic)

				f.SetColWidth(sheetName, "A", "H", 20)
			}
		}
	}

	// Save the file with updated data
	if err := f.SaveAs(filePath); err != nil {
		c.String(http.StatusInternalServerError, "Error menyimpan file: %v", err)
		return
	}
}
