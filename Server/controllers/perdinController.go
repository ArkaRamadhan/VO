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

func PerdinCreate(c *gin.Context) {
	// Mendapatkan data dari body request
	var requestBody perdinRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.Status(400)
		c.Error(err) // log the error
		return
	}

	// Menambahkan logging untuk melihat data yang diterima
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

	perdin := models.Perdin{
		NoPerdin:  requestBody.NoPerdin,
		Tanggal:   tanggal, // Gunakan tanggal yang telah diparsing, bisa jadi nil jika input kosong
		Hotel:     requestBody.Hotel,
		Transport: requestBody.Transport,
		CreateBy:  requestBody.CreateBy,
	}

	result := initializers.DB.Create(&perdin)

	if result.Error != nil {
		c.Status(400)
		return
	}

	// Mengembalikan hasil
	c.JSON(200, gin.H{
		"perdin": perdin,
	})
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

	initializers.DB.First(&perdin, id)

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

	if requestBody.Tanggal != nil {
		tanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(400, gin.H{"error": "Format tanggal tidak valid: " + err.Error()})
			return
		}
		perdin.Tanggal = &tanggal
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
	sheetNames := []string{"MEMO", "BERITA ACARA", "SK", "SURAT", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}
	// Create sheets and set headers for "SAG" only
	for _, sheetName := range sheetNames {
		if sheetName == "PERDIN" {
			f.NewSheet(sheetName)
			f.SetCellValue(sheetName, "A1", "No Perdin")
			f.SetCellValue(sheetName, "B1", "Tanggal")
			f.SetCellValue(sheetName, "C1", "Deskripsi")
			f.MergeCell(sheetName, "C1", "D1") // Menggabungkan sel C1 dan D1

			f.SetColWidth(sheetName, "A", "B", 20)
			f.SetColWidth(sheetName, "C", "D", 28)
			f.SetRowHeight(sheetName, 1, 28)
		} else {
			f.NewSheet(sheetName)
		}
	}

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

func UpdateSheetPerdin(c *gin.Context) {
	dir := ":\\excel"
	fileName := "its_report.xlsx"
	filePath := filepath.Join(dir, fileName)

	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		c.String(http.StatusBadRequest, "File tidak ada")
		return
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error membuka file: %v", err)
		return
	}
	defer f.Close()

	sheetName := "PERDIN"

	if _, err := f.GetSheetIndex(sheetName); err == nil {
		f.DeleteSheet(sheetName)
	}
	f.NewSheet(sheetName)

	f.SetCellValue(sheetName, "A1", "No Perdin")
	f.SetCellValue(sheetName, "B1", "Tanggal")
	f.SetCellValue(sheetName, "C1", "Deskripsi")
	f.MergeCell(sheetName, "C1", "D1") // Menggabungkan sel C1 dan D1

	// Mengatur lebar kolom
	f.SetColWidth(sheetName, "A", "D", 20)

	// Set text alignment to center for header cells
	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		c.String(http.StatusInternalServerError, "Error creating header style: %v", err)
		return
	}

	f.SetCellStyle(sheetName, "A1", "D1", headerStyle)

	var perdins []models.Perdin
	initializers.DB.Find(&perdins)

	for i, perdin := range perdins {
		rowNum := i + 2
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), perdin.NoPerdin)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), perdin.Tanggal.Format("2006-01-02"))
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), perdin.Hotel)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), perdin.Transport)
	}

	if err := f.SaveAs(filePath); err != nil {
		c.String(http.StatusInternalServerError, "Error menyimpan file: %v", err)
		return
	}

	c.Redirect(http.StatusFound, "/Perdin")
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

	log.Println("Processing rows...") // Log untuk memulai proses baris
	for i, row := range rows {
		if i == 0 { // Skip header or initial rows if necessary
			continue
		}

		// Count non-empty columns
		nonEmptyCount := 0
		for _, cell := range row {
			if cell != "" {
				nonEmptyCount++
			}
		}

		// Skip rows with less than 4 non-empty columns
		if nonEmptyCount < 2 {
			log.Printf("Row %d skipped: less than 4 columns filled, only %d filled", i+1, nonEmptyCount)
			continue
		}

		// Assign values from row, using nil if the cell is empty or column does not exist
		var (
			noPerdin  = getStringOrNil(getColumn(row, 0))
			tanggal   = getStringOrNil(getColumn(row, 1))
			hotel     = getStringOrNil(getColumn(row, 2))
			transport = getStringOrNil(getColumn(row, 3))
		)

		// Convert string dates to time.Time pointers if not nil
		var tanggalTime *time.Time
		if tanggal != nil {
			parsed, err := parseDate(*tanggal)
			if err != nil {
				log.Printf("Error parsing date from row %d: %v", i+1, err)
				continue
			}
			tanggalTime = &parsed
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
		log.Printf("Row %d imported successfully", i+1) // Log untuk setiap baris yang berhasil diimpor
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
