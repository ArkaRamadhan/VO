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

type SuratMasukRequest struct {
	ID         uint    `gorm:"primaryKey"`
	NoSurat    *string `json:"no_surat"`
	Title      *string `json:"title"`
	RelatedDiv *string `json:"related_div"`
	DestinyDiv *string `json:"destiny_div"`
	Tanggal    *string `json:"tanggal"`
	CreateBy   string  `json:"create_by"`
}

func UploadHandlerSuratMasuk(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/suratmasuk"
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

func GetFilesByIDSuratMasuk(c *gin.Context) {
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

func DeleteFileHandlerSuratMasuk(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/suratmasuk"
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

func DownloadFileHandlerSuratMasuk(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/suratmasuk"
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

func SuratMasukCreate(c *gin.Context) {
	// Get data off req body
	var requestBody SuratMasukRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}

	// Add some logging to see what's being received
	log.Println("Received request body:", requestBody)

	requestBody.CreateBy = c.MustGet("username").(string)

	var tanggal *time.Time // Deklarasi variabel tanggal sebagai pointer ke time.Time
	if requestBody.Tanggal != nil && *requestBody.Tanggal != "" {
		// Parse the date string only if it's not nil and not empty
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			log.Printf("Error parsing date: %v", err)
			c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
			return
		}
		tanggal = &parsedTanggal
	}

	surat_masuk := models.SuratMasuk{
		NoSurat:    requestBody.NoSurat,
		Title:      requestBody.Title,
		RelatedDiv: requestBody.RelatedDiv,
		DestinyDiv: requestBody.DestinyDiv,
		Tanggal:    tanggal, // Gunakan tanggal yang telah diparsing, bisa jadi nil jika input kosong
		CreateBy:   requestBody.CreateBy,
	}

	result := initializers.DB.Create(&surat_masuk)

	if result.Error != nil {
		c.Status(400)
		return
	}

	// Return it
	c.JSON(200, gin.H{
		"SuratMasuk": surat_masuk,
	})
}

func SuratMasukIndex(c *gin.Context) {

	// Get models from DB
	var surat_masuk []models.SuratMasuk
	initializers.DB.Find(&surat_masuk)

	//Respond with them
	c.JSON(200, gin.H{
		"SuratMasuk": surat_masuk,
	})
}

func SuratMasukShow(c *gin.Context) {

	id := c.Params.ByName("id")
	// Get models from DB
	var surat_masuk models.SuratMasuk

	initializers.DB.First(&surat_masuk, id)

	//Respond with them
	c.JSON(200, gin.H{
		"SuratMasuk": surat_masuk,
	})
}

func SuratMasukUpdate(c *gin.Context) {

	var requestBody SuratMasukRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}
	id := c.Params.ByName("id")

	var surat_masuk models.SuratMasuk
	initializers.DB.First(&surat_masuk, id)

	if err := initializers.DB.First(&surat_masuk, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "surat_masuk tidak ditemukan"})
		return
	}

	requestBody.CreateBy = c.MustGet("username").(string)
	surat_masuk.CreateBy = requestBody.CreateBy

	if requestBody.Tanggal != nil {
		tanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(400, gin.H{"error": "Format tanggal tidak valid: " + err.Error()})
			return
		}
		surat_masuk.Tanggal = &tanggal
	}

	if requestBody.NoSurat != nil {
		surat_masuk.NoSurat = requestBody.NoSurat
	} else {
		surat_masuk.NoSurat = surat_masuk.NoSurat // gunakan nilai yang ada dari database
	}

	if requestBody.Title != nil {
		surat_masuk.Title = requestBody.Title
	} else {
		surat_masuk.Title = surat_masuk.Title // gunakan nilai yang ada dari database
	}

	if requestBody.RelatedDiv != nil {
		surat_masuk.RelatedDiv = requestBody.RelatedDiv
	} else {
		surat_masuk.RelatedDiv = surat_masuk.RelatedDiv // gunakan nilai yang ada dari database
	}

	if requestBody.DestinyDiv != nil {
		surat_masuk.DestinyDiv = requestBody.DestinyDiv
	} else {
		surat_masuk.DestinyDiv = surat_masuk.DestinyDiv // gunakan nilai yang ada dari database
	}

	if requestBody.CreateBy != "" {
		surat_masuk.CreateBy = requestBody.CreateBy
	} else {
		surat_masuk.CreateBy = surat_masuk.CreateBy // gunakan nilai yang ada dari database
	}

	initializers.DB.Model(&surat_masuk).Updates(surat_masuk)

	c.JSON(200, gin.H{
		"surat_masuk": surat_masuk,
	})
}

func SuratMasukDelete(c *gin.Context) {

	//get id
	id := c.Params.ByName("id")

	// find the SuratMasuk
	var surat_masuk models.SuratMasuk

	if err := initializers.DB.First(&surat_masuk, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "surat masuk not found"})
		return
	}

	/// delete it
	if err := initializers.DB.Delete(&surat_masuk).Error; err != nil {
		c.JSON(404, gin.H{"error": "Surat Masuk Failed to Delete"})
		return
	}

	c.JSON(200, gin.H{
		"SuratMasuk": "Deleted",
	})
}

func CreateExcelSuratMasuk(c *gin.Context) {
	dir := "C:\\excel"
	baseFileName := "its_report_suratmasuk"
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
	sheetName := "SURAT MASUK"

	// Create sheets and set headers for "SURAT MASUK" only
	f.NewSheet(sheetName)
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
			Bold:   true,
			Size:   12,
			Color:  "FFFFFF",
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

// Fungsi untuk mengonversi serial Excel ke tanggal
func excelDateToTimeSuratMasuk(excelDate int) (time.Time, error) {
	// Excel menggunakan tanggal mulai 1 Januari 1900 (serial 1)
	baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	days := time.Duration(excelDate) * 24 * time.Hour
	return baseDate.Add(days), nil
}

func ImportExcelSuratMasuk(c *gin.Context) {
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
	sheetName := "SURAT MASUK"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Println("Processing rows...") // Log untuk memulai proses baris

	// Definisikan semua format tanggal yang mungkin
	dateFormats := []string{
		"2 January 2006",
		"02-Jan-06",
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
		"06-Jan-02",
		"02-Jan-06",
		"1-Jan-06",
		"06-Jan-02",
	}

	// Loop melalui baris dan simpan ke database
	for i, row := range rows {
		if i == 0 { // Lewati baris pertama yang merupakan header
			continue
		}
		if len(row) < 5 { // Pastikan ada cukup kolom
			log.Printf("Row %d skipped: less than 5 columns filled", i+1)
			continue
		}
		noSurat := row[0]
		title := row[1]
		related_div := row[2]
		destiny_div := row[3]
		tanggalString := row[4]

		var tanggal time.Time
		var parseErr error

		// Coba konversi dari serial Excel jika tanggalString adalah angka
		if serial, err := strconv.Atoi(tanggalString); err == nil {
			tanggal, parseErr = excelDateToTimeSuratMasuk(serial)
		} else {
			// Coba parse menggunakan format tanggal yang sudah ada
			for _, format := range dateFormats {
				tanggal, parseErr = time.Parse(format, tanggalString)
				if parseErr == nil {
					break // Keluar dari loop jika parsing berhasil
				}
			}
		}

		if parseErr != nil {
			log.Printf("Format tanggal tidak valid di baris %d: %v", i+1, parseErr)
			continue // Lewati baris ini jika format tanggal tidak valid
		}

		surat_masuk := models.SuratMasuk{
			NoSurat:    &noSurat,
			Title:      &title,
			RelatedDiv: &related_div,
			DestinyDiv: &destiny_div,
			Tanggal:    &tanggal,
			CreateBy:   c.MustGet("username").(string),
		}

		// Simpan ke database
		if err := initializers.DB.Create(&surat_masuk).Error; err != nil {
			log.Printf("Error saving record from row %d: %v", i+1, err)
			c.String(http.StatusInternalServerError, "Error saving record from row %d: %v", i+1, err)
			return
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data berhasil diimpor."})
}
