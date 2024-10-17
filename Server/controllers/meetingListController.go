package controllers

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"project-its/initializers"
	"project-its/models"
	"strconv"
	"time"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type MeetingListRequest struct {
	ID       uint    `gorm:"primaryKey"`
	Hari     *string `json:"hari"`
	Tanggal  *string `json:"tanggal"`
	Perihal  *string `json:"perihal"`
	Waktu    *string `json:"waktu"`
	Selesai  *string `json:"selesai"`
	Tempat   *string `json:"tempat"`
	Pic      *string `json:"pic"`
	Status   *string `json:"status"`
	CreateBy string  `json:"create_by"`
	Info     string  `json:"info"`
	Color    string  `json:"color"`
}

func UploadHandlerMeetingList(c *gin.Context) {
	id := c.PostForm("id")
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "File diperlukan"})
		return
	}

	// Konversi id dari string ke uint
	userID, err := strconv.ParseUint(id, 10, 32)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "ID tidak valid"})
		return
	}

	baseDir := "C:/UploadedFile/meetingschedule"
	dir := filepath.Join(baseDir, id)
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		os.MkdirAll(dir, 0755)
	}

	filePath := filepath.Join(dir, file.Filename)
	if err := c.SaveUploadedFile(file, filePath); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal menyimpan file"})
		return
	}

	// Menyimpan metadata file ke database
	newFile := models.File{
		UserID:      uint(userID), // Gunakan userID yang sudah dikonversi
		FilePath:    filePath,
		FileName:    file.Filename,
		ContentType: file.Header.Get("Content-Type"),
		Size:        file.Size,
	}
	result := initializers.DB.Create(&newFile)
	if result.Error != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal menyimpan metadata file"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "File berhasil diunggah"})
}

func GetFilesByIDMeetingList(c *gin.Context) {
	id := c.Param("id")

	var files []models.File
	result := initializers.DB.Where("user_id = ?", id).Find(&files)
	if result.Error != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal mengambil data file"})
		return
	}

	var fileNames []string
	for _, file := range files {
		fileNames = append(fileNames, file.FileName)
	}

	c.JSON(http.StatusOK, gin.H{"files": fileNames})
}

func DeleteFileHandlerMeetingList(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/meetingschedule"
	fullPath := filepath.Join(baseDir, id, filename)

	log.Printf("Attempting to delete file at path: %s", fullPath)

	// Hapus file dari sistem file
	err = os.Remove(fullPath)
	if err != nil {
		log.Printf("Error deleting file: %v", err)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to delete file"})
		return
	}

	// Hapus metadata file dari database
	result := initializers.DB.Where("file_path = ?", fullPath).Delete(&models.File{})
	if result.Error != nil {
		log.Printf("Error deleting file metadata: %v", result.Error)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to delete file metadata"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "File deleted successfully"})
}

func DownloadFileHandlerMeetingList(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/meetingschedule"
	fullPath := filepath.Join(baseDir, id, filename)

	log.Printf("Full path for download: %s", fullPath)

	// Periksa keberadaan file di database
	var file models.File
	result := initializers.DB.Where("file_path = ?", fullPath).First(&file)
	if result.Error != nil {
		log.Printf("File not found in database: %v", result.Error)
		c.JSON(http.StatusNotFound, gin.H{"error": "File tidak ditemukan"})
		return
	}

	// Periksa keberadaan file di sistem file
	if _, err := os.Stat(fullPath); os.IsNotExist(err) {
		log.Printf("File not found in system: %s", fullPath)
		c.JSON(http.StatusNotFound, gin.H{"error": "File tidak ditemukan di sistem file"})
		return
	}

	log.Printf("File downloaded successfully: %s", fullPath)
	c.File(fullPath)
}

func MeetingListIndex(c *gin.Context) {

	var meetingList []models.MeetingSchedule

	initializers.DB.Find(&meetingList)

	c.JSON(200, gin.H{
		"meetingschedule": meetingList,
	})

}

func MeetingListCreate(c *gin.Context) {

	var requestBody MeetingListRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}

	// Add some logging to see what's being received
	log.Println("Received request body:", requestBody)

	// Parse the date string
	tanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
	if err != nil {
		log.Printf("Error parsing date: %v", err)
		c.Status(400)
		c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
		return
	}

	requestBody.CreateBy = c.MustGet("username").(string)

	meetingList := models.MeetingSchedule{
		Hari:     requestBody.Hari,
		Tanggal:  &tanggal,
		Perihal:  requestBody.Perihal,
		Waktu:    requestBody.Waktu,
		Selesai:  requestBody.Selesai,
		Tempat:   requestBody.Tempat,
		Pic:      requestBody.Pic,
		Status:   requestBody.Status,
		CreateBy: requestBody.CreateBy,
		Color:    requestBody.Color,
	}

	result := initializers.DB.Create(&meetingList)

	if result.Error != nil {
		c.Status(400)
		c.JSON(400, gin.H{"error": "Failed to create Meeting: " + result.Error.Error()})
		return
	}

	c.JSON(201, gin.H{
		"meetingschedule": meetingList,
	})

}

func MeetingListShow(c *gin.Context) {

	id := c.Params.ByName("id")

	var meetingList models.MeetingSchedule

	if err := initializers.DB.First(&meetingList, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "Meeting not found"}) // Tambahkan log ini
		return
	}

	c.JSON(200, gin.H{
		"meetingschedule": meetingList,
	})

}

func MeetingListUpdate(c *gin.Context) {

	var requestBody MeetingListRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(400, gin.H{"error": "Invalid request body: " + err.Error()})
		return
	}

	id := c.Params.ByName("id")

	var meetingList models.MeetingSchedule

	if err := initializers.DB.First(&meetingList, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "meeting not found"})
		return
	}

	requestBody.CreateBy = c.MustGet("username").(string)
	meetingList.CreateBy = requestBody.CreateBy

	if requestBody.Tanggal != nil {
		tanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(400, gin.H{"error": "Format tanggal tidak valid: " + err.Error()})
			return
		}
		meetingList.Tanggal = &tanggal
	}

	if requestBody.Hari != nil {
		meetingList.Hari = requestBody.Hari
	} else {
		meetingList.Hari = meetingList.Hari
	}

	if requestBody.Perihal != nil {
		meetingList.Perihal = requestBody.Perihal
	} else {
		meetingList.Perihal = meetingList.Perihal

	}

	if requestBody.Waktu != nil {
		meetingList.Waktu = requestBody.Waktu
	} else {
		meetingList.Waktu = meetingList.Waktu
	}

	if requestBody.Selesai != nil {
		meetingList.Selesai = requestBody.Selesai
	} else {
		meetingList.Selesai = meetingList.Selesai
	}

	if requestBody.Tempat != nil {
		meetingList.Tempat = requestBody.Tempat
	} else {
		meetingList.Tempat = meetingList.Tempat
	}

	if requestBody.Pic != nil {
		meetingList.Pic = requestBody.Pic
	} else {
		meetingList.Pic = meetingList.Pic
	}

	if requestBody.Status != nil {
		meetingList.Status = requestBody.Status
	} else {
		meetingList.Status = meetingList.Status
	}

	if requestBody.Color != "" {
		meetingList.Color = requestBody.Color
	} else {
		meetingList.Color = meetingList.Color
	}

	initializers.DB.Save(&meetingList)

	c.JSON(200, gin.H{
		"meetinglist": meetingList,
	})
}

func MeetingListDelete(c *gin.Context) {

	id := c.Params.ByName("id")

	var meetingList models.MeetingSchedule

	if err := initializers.DB.First(&meetingList, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "meeting not found"})
		return
	}

	if err := initializers.DB.Delete(&meetingList).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete Memo: " + err.Error()})
		return
	}

	c.Status(204)

}

func CreateExcelMeetingList(c *gin.Context) {
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
	sheetNames := []string{"MEMO", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}

	// Create sheets and set headers for "MEETING SCHEDULE" only
	for _, sheetName := range sheetNames {
		if sheetName == "MEETING SCHEDULE" {
			f.NewSheet(sheetName)
			f.SetCellValue(sheetName, "A1", "Hari")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Perihal")
			f.SetCellValue(sheetName, "D1", "Waktu")
			f.SetCellValue(sheetName, "E1", "Selesai")
			f.SetCellValue(sheetName, "F1", "Tempat")
			f.SetCellValue(sheetName, "G1", "Status")
			f.SetCellValue(sheetName, "H1", "PIC")

			f.SetColWidth(sheetName, "A", "H", 20)
			f.SetRowHeight(sheetName, 1, 20)
		} else {
			f.NewSheet(sheetName)
		}
	}

	styleHeader, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#6EB6F8"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
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

	err = f.SetCellStyle("MEETING SCHEDULE", "A1", "H1", styleHeader)

	// Fetch initial data from the database
	var meetingList []models.MeetingSchedule
	initializers.DB.Find(&meetingList)

	// Definisikan gaya untuk border
	styleAll, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		c.String(http.StatusInternalServerError, "Error membuat gaya: %v", err)
		return
	}

	// Definisikan gaya untuk status yang berbeda
	styleDone, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000", Bold: true},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#5cb85c"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	styleCancel, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000", Bold: true},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#d9534f"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	styleReschedule, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000", Bold: true},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#0275d8"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	styleOnProgress, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000", Bold: true},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#f0ad4e"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	// Tulis data awal ke lembar "MEETING SCHEDULE"
	meetingListSheetName := "MEETING SCHEDULE"
	for i, meetingList := range meetingList {
		rowNum := i + 2 // Start from the second row (first row is header)
		f.SetCellValue(meetingListSheetName, fmt.Sprintf("A%d", rowNum), *meetingList.Hari)
		f.SetCellValue(meetingListSheetName, fmt.Sprintf("B%d", rowNum), meetingList.Tanggal.Format("2006-01-02"))
		f.SetCellValue(meetingListSheetName, fmt.Sprintf("C%d", rowNum), *meetingList.Perihal)

		// Handle Waktu
		if meetingList.Waktu != nil {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("D%d", rowNum), *meetingList.Waktu)
		} else {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("D%d", rowNum), "")
		}

		// Handle Selesai
		if meetingList.Selesai != nil {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("E%d", rowNum), *meetingList.Selesai)
		} else {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("E%d", rowNum), "")
		}

		if meetingList.Tempat != nil {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("F%d", rowNum), *meetingList.Tempat)
		} else {
			f.SetCellValue(meetingListSheetName, fmt.Sprintf("F%d", rowNum), "")
		}

		f.SetCellValue(meetingListSheetName, fmt.Sprintf("G%d", rowNum), *meetingList.Status)
		f.SetCellValue(meetingListSheetName, fmt.Sprintf("H%d", rowNum), *meetingList.Pic)

		// Terapkan gaya border untuk semua sel
		for col := 'A'; col <= 'H'; col++ {
			cellName := fmt.Sprintf("%c%d", col, rowNum)
			f.SetCellStyle(meetingListSheetName, cellName, cellName, styleAll)
		}

		// Terapkan gaya khusus untuk status
		switch *meetingList.Status {
		case "Done":
			f.SetCellStyle(meetingListSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), styleDone)
		case "Cancel":
			f.SetCellStyle(meetingListSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), styleCancel)
		case "Reschedule":
			f.SetCellStyle(meetingListSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), styleReschedule)
		case "On Progress":
			f.SetCellStyle(meetingListSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), styleOnProgress)
		}
	}

	// Delete the default "Sheet1" sheet
	if err := f.DeleteSheet("Sheet1"); err != nil {
		panic(err) // Handle error jika bukan error "sheet tidak ditemukan"
	}

	// Save the newly created file
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

func UpdateSheetMeetingList(c *gin.Context) {
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

	// Define sheet name
	sheetName := "MEETING SCHEDULE"

	// Check if the sheet exists
	sheetIndex, err := f.GetSheetIndex(sheetName)
	if err != nil || sheetIndex == -1 {
		c.String(http.StatusBadRequest, "Lembar kerja MEETING SCHEDULE tidak ditemukan")
		return
	}

	// Clear existing data except the header by deleting rows
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}
	for i := 1; i < len(rows); i++ { // Start from 1 to keep the header
		f.RemoveRow(sheetName, 2) // Always remove the second row since the sheet compresses up
	}

	// Fetch updated data from the database
	var meetingList []models.MeetingSchedule
	initializers.DB.Find(&meetingList)

	// Write data rows
	for i, meetingList := range meetingList {
		rowNum := i + 2 // Start from the second row (first row is header)
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
	}

	// Save the file with updated data
	if err := f.SaveAs(filePath); err != nil {
		c.String(http.StatusInternalServerError, "Error menyimpan file: %v", err)
		return
	}

}

func ImportExcelMeetingList(c *gin.Context) {
	// Mengambil file dari form upload
	file, _, err := c.Request.FormFile("file")
	if err != nil {
		c.String(http.StatusBadRequest, "Error retrieving the file: %v", err)
		return
	}
	defer file.Close()

	// Simpan file sementara jika perlu
	tempFile, err := os.CreateTemp("", "*.xlsx")
	if err != nil {
		c.String(http.StatusInternalServerError, "Error creating temporary file: %v", err)
		return
	}
	defer os.Remove(tempFile.Name()) // Hapus file sementara setelah selesai

	// Salin file dari request ke file sementara
	if _, err := file.Seek(0, 0); err != nil {
		c.String(http.StatusInternalServerError, "Error seeking file: %v", err)
		return
	}
	if _, err := io.Copy(tempFile, file); err != nil {
		c.String(http.StatusInternalServerError, "Error copying file: %v", err)
		return
	}

	// Buka file Excel dari file sementara
	tempFile.Seek(0, 0) // Reset pointer ke awal file
	f, err := excelize.OpenFile(tempFile.Name())
	if err != nil {
		c.String(http.StatusInternalServerError, "Error opening file: %v", err)
		return
	}
	defer f.Close()

	// Pilih sheet
	sheetName := "MEETING SCHEDULE"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	// Loop melalui baris dan simpan ke database
	for i, row := range rows {
		if i == 0 {
			// Lewati header baris jika ada
			continue
		}
		if len(row) < 4 {
			// Pastikan ada cukup kolom
			continue
		}
		hari := row[0]
		tanggal := row[1]
		perihal := row[2]
		waktu := row[3]
		selesai := row[4]
		tempat := row[5]
		status := row[6]
		pic := row[7]

		// Parse tanggal
		tanggalString, err := time.Parse("2006-01-02", tanggal)
		if err != nil {
			c.String(http.StatusBadRequest, "Invalid date format in row %d: %v", i+1, err)
			return
		}

		meetingList := models.MeetingSchedule{
			Hari:     &hari,
			Tanggal:  &tanggalString,
			Perihal:  &perihal,
			Waktu:    &waktu,
			Selesai:  &selesai,
			Tempat:   &tempat,
			Status:   &status,
			Pic:      &pic,
			CreateBy: c.MustGet("username").(string),
		}

		// Simpan ke database
		if err := initializers.DB.Create(&meetingList).Error; err != nil {
			log.Printf("Error saving record from row %d: %v", i+1, err)
			c.String(http.StatusInternalServerError, "Error saving record from row %d: %v", i+1, err)
			return
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully."})
}
