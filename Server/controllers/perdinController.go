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

type perdinRequest struct {
	ID        uint    `gorm:"primaryKey"`
	NoPerdin  *string `json:"no_perdin"`
	Tanggal   *string `json:"tanggal"`
	Hotel     *string `json:"hotel"`
	Transport *string `json:"transport"`
	CreateBy  string  `json:"create_by"`
}

func UploadHandlerPerdin(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/perdin"
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

func GetFilesByIDPerdin(c *gin.Context) {
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

func DeleteFileHandlerPerdin(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/perdin"
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

func DownloadFileHandlerPerdin(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/perdin"
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

func GetLatestPerdinNumber(NoPerdin string) (string, error) {
	var lastPerdin models.Perdin
	// Ubah pencarian untuk menggunakan format yang benar
	searchPattern := fmt.Sprintf("%%/%s/%%", NoPerdin) // Ini akan mencari format seperti '%/ITS-SAG/M/%'
	if err := initializers.DB.Where("no_perdin LIKE ?", searchPattern).Order("id desc").First(&lastPerdin).Error; err != nil {
		if errors.Is(err, gorm.ErrRecordNotFound) {
			return "00001", nil // Jika tidak ada catatan, kembalikan 00001
		}
		return "", err
	}

	// Ambil nomor memo terakhir, pisahkan, dan tambahkan 1
	parts := strings.Split(*lastPerdin.NoPerdin, "/")
	if len(parts) > 0 {
		number, err := strconv.Atoi(parts[0])
		if err != nil {
			return "", err
		}
		return fmt.Sprintf("%05d", number+1), nil // Tambahkan 1 ke nomor terakhir
	}

	return "00001", nil
}

func PerdinCreate(c *gin.Context) {
	var requestBody perdinRequest
	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request"})
		return
	}

	// Check if NoPerdin is empty and set default
	if requestBody.NoPerdin == nil || *requestBody.NoPerdin == "" {
		nomor, err := GetLatestPerdinNumber("PD-ITS")
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to generate perdin number"})
			return
		}
		tahun := time.Now().Year()
		noPerdin := fmt.Sprintf("%s/PD-ITS/%d", nomor, tahun)
		requestBody.NoPerdin = &noPerdin
	}

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

	requestBody.CreateBy = c.MustGet("username").(string)

	// Create perdin with the requestBody data
	perdin := models.Perdin{
		NoPerdin:  requestBody.NoPerdin,
		Tanggal:   tanggal,
		Hotel:     requestBody.Hotel,
		Transport: requestBody.Transport,
	}

	if result := initializers.DB.Create(&perdin); result.Error != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create perdin"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"perdin": perdin})
}

func PerdinIndex(c *gin.Context) {

	// Get models from DB
	var perdin []models.Perdin
	initializers.DB.Find(&perdin)

	//Respond with them
	c.JSON(200, gin.H{
		"perdin": perdin,
	})
}

func PerdinShow(c *gin.Context) {

	id := c.Params.ByName("id")
	// Get models from DB
	var perdin models.Perdin

	if err := initializers.DB.First(&perdin, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "perdin tidak ditemukan"})
		return
	}

	// Log field yang terambil
	log.Printf("Perdin retrieved: ID=%d, NoPerdin=%s, Tanggal=%v, Hotel=%s, Transport=%s, CreateBy=%s",
		perdin.ID, getStringValue(perdin.NoPerdin), perdin.Tanggal, getStringValue(perdin.Hotel), getStringValue(perdin.Transport), perdin.CreateBy)

	//Respond with them
	c.JSON(200, gin.H{
		"perdin": perdin,
	})
}

func PerdinUpdate(c *gin.Context) {

	var requestBody perdinRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}

	id := c.Params.ByName("id")

	var perdin models.Perdin
	initializers.DB.First(&perdin, id)

	if err := initializers.DB.First(&perdin, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "perdin tidak ditemukan"})
		return
	}

	requestBody.CreateBy = c.MustGet("username").(string)
	perdin.CreateBy = requestBody.CreateBy

	nomor, err := GetLatestPerdinNumber(*requestBody.NoPerdin)
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
	if *requestBody.NoPerdin == "PD-ITS" {
		noPerdin := fmt.Sprintf("%s/PD-ITS/%d", nomor, tahun)
		requestBody.NoPerdin = &noPerdin
		log.Printf("Generated NoPerdin for Perdin: %s", *requestBody.NoPerdin) // Log nomor memo
	}

	// Update tanggal jika diberikan dan tidak kosong
	if requestBody.Tanggal != nil && *requestBody.Tanggal != "" {
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		perdin.Tanggal = &parsedTanggal
	}

	if requestBody.NoPerdin != nil {
		perdin.NoPerdin = requestBody.NoPerdin
	} else {
		perdin.NoPerdin = perdin.NoPerdin // gunakan nilai yang ada dari database
	}

	if requestBody.Transport != nil {
		perdin.Transport = requestBody.Transport
	} else {
		perdin.Transport = perdin.Transport // gunakan nilai yang ada dari database
	}

	if requestBody.Hotel != nil {
		perdin.Hotel = requestBody.Hotel
	} else {
		perdin.Hotel = perdin.Hotel // gunakan nilai yang ada dari database
	}

	if requestBody.CreateBy != "" {
		perdin.CreateBy = requestBody.CreateBy
	} else {
		perdin.CreateBy = perdin.CreateBy // gunakan nilai yang ada dari database
	}

	initializers.DB.Model(&perdin).Updates(perdin)

	c.JSON(200, gin.H{
		"perdin": perdin,
	})

}

func PerdinDelete(c *gin.Context) {

	//get id
	id := c.Params.ByName("id")

	// find the perdin
	var perdin models.Perdin

	if err := initializers.DB.First(&perdin, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "Perdin not found"})
		return
	}

	/// delete it
	if err := initializers.DB.Delete(&perdin).Error; err != nil {
		c.JSON(404, gin.H{"error": "Perdin Failed to Delete"})
		return
	}

	c.JSON(200, gin.H{
		"perdin": "Perdin deleted",
	})
}

func CreateExcelPerdin(c *gin.Context) {
	dir := ":\\excel"
	baseFileName := "its_report_perdin"
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
	sheetName := "PERDIN"

	f.NewSheet(sheetName)
	f.SetCellValue(sheetName, "A1", "No Perdin")
	f.SetCellValue(sheetName, "B1", "Tanggal")
	f.SetCellValue(sheetName, "C1", "Deskripsi")
	f.MergeCell(sheetName, "C1", "D1") // Menggabungkan sel C1 dan D1

	f.SetColWidth(sheetName, "A", "B", 20)
	f.SetColWidth(sheetName, "C", "D", 28)
	f.SetRowHeight(sheetName, 1, 28)

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

func excelDateToTimePerdin(excelDate int) (time.Time, error) {
	// Excel menggunakan tanggal mulai 1 Januari 1900 (serial 1)
	baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	days := time.Duration(excelDate) * 24 * time.Hour
	return baseDate.Add(days), nil
}

func ImportExcelPerdin(c *gin.Context) {
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

	sheetName := "PERDIN"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Println("Processing rows...")

	// Definisikan semua format tanggal yang mungkin
	dateFormats := []string{
		"02-Jan-06",
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
		"06-Jan-02",
	}

	for i, row := range rows {
		if i == 0 { // Lewati baris pertama yang merupakan header
			continue
		}
		if len(row) < 2 { // Pastikan ada cukup kolom
			log.Printf("Row %d skipped: less than 2 columns filled", i+1)
			continue
		}

		// Ambil data dari kolom
		noPerdin := getStringOrNil(getColumn(row, 0))
		tanggalStr := getStringOrNil(getColumn(row, 1))
		hotel := getStringOrNil(getColumn(row, 2))
		transport := getStringOrNil(getColumn(row, 3))

		var tanggalTime *time.Time
		if tanggalStr != nil {
			var parseErr error
			// Coba konversi dari serial Excel jika tanggalStr adalah angka
			if serial, err := strconv.Atoi(*tanggalStr); err == nil {
				parsed, parseErr := excelDateToTimePerdin(serial)
				if parseErr == nil {
					tanggalTime = &parsed
				}
			} else {
				// Coba parse menggunakan format tanggal yang sudah ada
				for _, format := range dateFormats {
					parsed, parseErr := time.Parse(format, *tanggalStr)
					if parseErr == nil {
						tanggalTime = &parsed
						break // Keluar dari loop jika parsing berhasil
					}
				}
			}

			if parseErr != nil {
				log.Printf("Format tanggal tidak valid di baris %d: %v", i+1, parseErr)
				continue // Lewati baris ini jika format tanggal tidak valid
			}
		}

		perdin := models.Perdin{
			Tanggal:   tanggalTime,
			NoPerdin:  noPerdin,
			Hotel:     hotel,
			Transport: transport,
			CreateBy:  c.MustGet("username").(string),
		}

		if err := initializers.DB.Create(&perdin).Error; err != nil {
			log.Printf("Error saving record from row %d: %v", i+1, err)
			continue
		}
		log.Printf("Row %d imported successfully", i+1)
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}

// Helper function to get the string value from a pointer
func getStringValue(str *string) string {
	if str == nil {
		return ""
	}
	return *str
}
