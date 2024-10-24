package controllers

import (
	"errors"
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
	"strings"
	"time"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
	"gorm.io/gorm"
)

type MemoRequest struct {
	ID       uint    `gorm:"primaryKey"`
	Tanggal  *string `json:"tanggal"`
	NoMemo   *string `json:"no_memo"`
	Perihal  *string `json:"perihal"`
	Pic      *string `json:"pic"`
	Kategori *string `json:"kategori"`
	CreateBy string  `json:"create_by"`
}

func UploadHandlerMemo(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/memo"
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

func GetFilesByIDMemo(c *gin.Context) {
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

func DeleteFileHandlerMemo(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/memo"
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

func DownloadFileHandlerMemo(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/memo"
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

func GetLatestMemoNumber(category string) (string, error) {
	var lastMemo models.Memo
	searchPattern := fmt.Sprintf("%%/%s/M/%%", category) // Sesuaikan pencarian berdasarkan kategori
	if err := initializers.DB.Where("no_memo LIKE ?", searchPattern).Order("id desc").First(&lastMemo).Error; err != nil {
		if errors.Is(err, gorm.ErrRecordNotFound) {
			// Jika tidak ada memo yang ditemukan untuk kategori ini, kembalikan "00001"
			return "00001", nil
		}
		return "", err
	}

	// Ekstrak nomor terakhir dan tambahkan 1
	parts := strings.Split(*lastMemo.NoMemo, "/")
	if len(parts) > 0 {
		number, err := strconv.Atoi(parts[0])
		if err != nil {
			return "", err
		}
		return fmt.Sprintf("%05d", number+1), nil
	}

	return "00001", nil
}

func MemoIndex(c *gin.Context) {

	var memosag []models.Memo

	initializers.DB.Find(&memosag)

	c.JSON(200, gin.H{
		"memo": memosag,
	})

}

func MemoCreate(c *gin.Context) {
	var requestBody MemoRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}

	log.Println("Received request body:", requestBody)

	var tanggal *time.Time
	if requestBody.Tanggal != nil && *requestBody.Tanggal != "" {
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			log.Printf("Error parsing date: %v", err)
			c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
			return
		}
		tanggal = &parsedTanggal
	}

	log.Printf("Parsed date: %v", tanggal) // Tambahkan log ini untuk melihat tanggal yang diparsing

	nomor, err := GetLatestMemoNumber(*requestBody.NoMemo)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest memo number"})
		return
	}

	// Cek apakah nomor yang diterima adalah "00001"
	if nomor == "00001" {
		// Jika "00001", berarti ini adalah entri pertama
		log.Println("This is the first memo entry.")
	}

	tahun := time.Now().Year()
	// Menentukan format NoMemo berdasarkan kategori
	if *requestBody.NoMemo == "ITS-SAG" {
		noMemo := fmt.Sprintf("%s/ITS-SAG/M/%d", nomor, tahun)
		requestBody.NoMemo = &noMemo
		log.Printf("Generated NoMemo for ITS-SAG: %s", *requestBody.NoMemo) // Log nomor memo
	} else if *requestBody.NoMemo == "ITS-ISO" {
		noMemo := fmt.Sprintf("%s/ITS-ISO/M/%d", nomor, tahun)
		requestBody.NoMemo = &noMemo
		log.Printf("Generated NoMemo for ITS-ISO: %s", *requestBody.NoMemo) // Log nomor memo
	}

	requestBody.CreateBy = c.MustGet("username").(string)

	memosag := models.Memo{
		Tanggal:  tanggal,
		NoMemo:   requestBody.NoMemo, // Menggunakan NoMemo yang sudah diformat
		Perihal:  requestBody.Perihal,
		Pic:      requestBody.Pic,
		CreateBy: requestBody.CreateBy,
	}

	result := initializers.DB.Create(&memosag)
	if result.Error != nil {
		log.Printf("Error saving memo: %v", result.Error)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create Memo Sag"})
		return
	}
	log.Printf("Memo created successfully: %v", memosag)

	c.JSON(201, gin.H{
		"memo": memosag,
	})
}

func MemoShow(c *gin.Context) {

	id := c.Params.ByName("id")

	var memosag models.Memo

	initializers.DB.First(&memosag, id)

	// Log field yang terambil
	log.Printf("Memo retrieved: ID=%d, NoMemo=%s, Tanggal=%v, Perihal=%s, Pic=%s, CreateBy=%s",
		memosag.ID, getStringValue(memosag.NoMemo), memosag.Tanggal, getStringValue(memosag.Perihal), getStringValue(memosag.Pic), memosag.CreateBy)

	c.JSON(200, gin.H{
		"memo": memosag,
	})

}

func MemoUpdate(c *gin.Context) {
	var requestBody MemoRequest

	// Bind JSON request body
	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request body: " + err.Error()})
		return
	}

	// Ambil ID memo dari parameter URL
	id := c.Param("id")
	var memo models.Memo

	// Cari memo berdasarkan ID
	if err := initializers.DB.First(&memo, id).Error; err != nil {
		log.Printf("Memo with ID %s not found: %v", id, err)
		c.JSON(http.StatusNotFound, gin.H{"error": "Memo not found"})
		return
	}

	// Cek apakah NoMemo akan di-update
	if requestBody.NoMemo != nil && *requestBody.NoMemo != "" && *memo.NoMemo != *requestBody.NoMemo {
		nomor, err := GetLatestMemoNumber(*requestBody.NoMemo)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest memo number"})
			return
		}

		tahun := time.Now().Year()

		// Menentukan format NoMemo berdasarkan kategori
		if *requestBody.NoMemo == "ITS-SAG" {
			noMemo := fmt.Sprintf("%s/ITS-SAG/M/%d", nomor, tahun)
			requestBody.NoMemo = &noMemo
			log.Printf("Generated NoMemo for ITS-SAG: %s", *requestBody.NoMemo)
		} else if *requestBody.NoMemo == "ITS-ISO" {
			noMemo := fmt.Sprintf("%s/ITS-ISO/M/%d", nomor, tahun)
			requestBody.NoMemo = &noMemo
			log.Printf("Generated NoMemo for ITS-ISO: %s", *requestBody.NoMemo)
		}

		// Update NoMemo hanya jika ada perubahan
		memo.NoMemo = requestBody.NoMemo
	}

	// Update tanggal jika diberikan dan tidak kosong
	if requestBody.Tanggal != nil && *requestBody.Tanggal != "" {
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		memo.Tanggal = &parsedTanggal
	}

	// Update Perihal jika ada
	if requestBody.Perihal != nil {
		memo.Perihal = requestBody.Perihal
	}

	// Update PIC jika ada
	if requestBody.Pic != nil {
		memo.Pic = requestBody.Pic
	}

	// Simpan perubahan
	if err := initializers.DB.Save(&memo).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to update memo"})
		return
	}

	// Return response sukses
	c.JSON(http.StatusOK, gin.H{
		"message": "Memo updated successfully",
		"memo":    memo,
	})
}

func MemoDelete(c *gin.Context) {

	id := c.Params.ByName("id")

	var memosag models.Memo

	if err := initializers.DB.First(&memosag, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Memo not found"})
		return
	}

	if err := initializers.DB.Delete(&memosag).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete Memo: " + err.Error()})
		return
	}

	c.Status(204)

}

func exportMemoToExcel(memos []models.Memo) (*excelize.File, error) {
	// Buat file Excel baru
	f := excelize.NewFile()

	sheetName := "MEMO"
	f.NewSheet(sheetName)
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

	f.DeleteSheet("Sheet1")

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

	return f, nil
}

// Handler untuk melakukan export Excel dengan Gin
func ExportMemoHandler(c *gin.Context) {
	// Data memo contoh
	var memos []models.Memo
	initializers.DB.Find(&memos)

	// Buat file Excel
	f, err := exportMemoToExcel(memos)
	if err != nil {
		c.String(http.StatusInternalServerError, "Gagal mengekspor data ke Excel")
		return
	}

	// Set nama file dan header untuk download
	fileName := fmt.Sprintf("its_report_memo.xlsx")
	c.Header("Content-Disposition", "attachment; filename="+fileName)
	c.Header("Content-Type", "application/octet-stream")

	// Simpan file Excel ke dalam buffer
	if err := f.Write(c.Writer); err != nil {
		c.String(http.StatusInternalServerError, "Gagal menyimpan file Excel")
	}
}

func excelDateToTimeMemo(excelDate int) (time.Time, error) {
	baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	days := time.Duration(excelDate) * 24 * time.Hour
	return baseDate.Add(days), nil
}

func ImportExcelMemo(c *gin.Context) {
	file, _, err := c.Request.FormFile("file")
	if err != nil {
		c.String(http.StatusBadRequest, "Error retrieving the file: %v", err)
		return
	}
	defer file.Close()

	tempFile, err := os.CreateTemp("", "*.xlsx")
	if err != nil {
		c.String(http.StatusInternalServerError, "Error creating temporary file: %v", err)
		return
	}
	defer os.Remove(tempFile.Name())

	if _, err := io.Copy(tempFile, file); err != nil {
		c.String(http.StatusInternalServerError, "Error copying file: %v", err)
		return
	}

	tempFile.Seek(0, 0)
	f, err := excelize.OpenFile(tempFile.Name())
	if err != nil {
		c.String(http.StatusInternalServerError, "Error opening file: %v", err)
		return
	}
	defer f.Close()

	sheetName := "MEMO"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Println("Processing rows...")

	// Definisikan semua format tanggal yang mungkin
	dateFormats := []string{
		"02-Jan-06",
		"02 Jan 06",
		"02-01-06",
		"02/01/2006",
		"01-02-2006",
		"06-Jan-02",
		"2 January 2006",
		"2006-01-02",
		"02-01-2006",
		"01/02/2006",
		"2006.01.02",
		"02/01/2006",
		"Jan 2, 06",
		"Jan 2, 2006",
		"01/02/06",
		"02/01/06",
		"06/02/01",
		"06/01/02",
		"1-Jan-06",
	}

	for i, row := range rows {
		if i == 0 { // Lewati baris pertama yang merupakan header
			continue
		}
		// Pastikan minimal 4 kolom terisi
		nonEmptyColumns := 0
		for _, col := range row {
			if col != "" {
				nonEmptyColumns++
			}
		}
		if nonEmptyColumns < 2 {
			log.Printf("Baris %d dilewati: hanya %d kolom terisi", i+1, nonEmptyColumns)
			continue
		}

		// Ambil data dari kolom dengan penanganan jika kolom kosong
		tanggalStr := ""
		if len(row) > 0 {
			tanggalStr = row[0]
		}
		noMemo := ""
		if len(row) > 1 {
			noMemo = row[1]
		}
		perihal := ""
		if len(row) > 2 {
			perihal = row[2]
		}
		pic := ""
		if len(row) > 3 {
			pic = row[3]
		}

		var tanggal *time.Time
		var parseErr error
		if tanggalStr != "" {
			// Coba parse menggunakan format tanggal yang sudah ada
			for _, format := range dateFormats {
				var parsedTanggal time.Time
				parsedTanggal, parseErr = time.Parse(format, tanggalStr)
				if parseErr == nil {
					tanggal = &parsedTanggal
					break // Keluar dari loop jika parsing berhasil
				}
			}
			if parseErr != nil {
				log.Printf("Format tanggal tidak valid di baris %d: %v", i+1, parseErr)
			}
		}

		memo := models.Memo{
			Tanggal:  tanggal,
			NoMemo:   &noMemo,
			Perihal:  &perihal,
			Pic:      &pic,
			CreateBy: c.MustGet("username").(string),
		}

		if err := initializers.DB.Create(&memo).Error; err != nil {
			log.Printf("Error saving memo record from row %d: %v", i+1, err)
		} else {
			log.Printf("Memo Row %d imported successfully", i+1)
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}
