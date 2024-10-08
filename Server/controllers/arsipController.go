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

type arsipRequest struct {
	ID                uint    `gorm:"primaryKey"`
	NoArsip           *string `json:"no_arsip"`
	JenisDokumen      *string `json:"jenis_dokumen"`
	NoDokumen         *string `json:"no_dokumen"`
	Perihal           *string `json:"perihal"`
	NoBox             *string `json:"no_box"`
	TanggalDokumen    *string `json:"tanggal_dokumen"`
	TanggalPenyerahan *string `json:"tanggal_penyerahan"`
	Keterangan        *string `json:"keterangan"`
	CreateBy          string  `json:"create_by"`
}

func UploadHandlerArsip(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/arsip"
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

func GetFilesByIDArsip(c *gin.Context) {
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

func DeleteFileHandlerArsip(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/arsip"
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

func DownloadFileHandlerArsip(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/arsip"
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

func ArsipIndex(c *gin.Context) {
	var arsip []models.Arsip
	initializers.DB.Find(&arsip)
	c.JSON(200, gin.H{
		"arsip": arsip,
	})
}

// Fungsi untuk membuat arsip baru
func ArsipCreate(c *gin.Context) {
	var requestBody arsipRequest
	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request data"})
		return
	}
	requestBody.CreateBy = c.MustGet("username").(string)

	var tanggal *time.Time
	if requestBody.TanggalDokumen != nil && *requestBody.TanggalDokumen != "" {
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.TanggalDokumen)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		tanggal = &parsedTanggal
	}

	arsip := models.Arsip{
		NoArsip:           requestBody.NoArsip,
		JenisDokumen:      requestBody.JenisDokumen,
		NoDokumen:         requestBody.NoDokumen,
		Perihal:           requestBody.Perihal,
		NoBox:             requestBody.NoBox,
		Keterangan:        requestBody.Keterangan,
		TanggalDokumen:    tanggal,
		TanggalPenyerahan: tanggal, // Assuming same date handling for TanggalPenyerahan
		CreateBy:          requestBody.CreateBy,
	}

	if err := initializers.DB.Create(&arsip).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create arsip"})
		return
	}
	c.JSON(http.StatusOK, gin.H{"arsip": arsip})
}

func ArsipShow(c *gin.Context) {
	id := c.Param("id")
	var arsip models.Arsip
	if err := initializers.DB.First(&arsip, id).Error; err != nil {
		c.JSON(http.StatusNotFound, gin.H{"error": "Arsip not found"})
		return
	}
	c.JSON(http.StatusOK, gin.H{"arsip": arsip})
}

func ArsipUpdate(c *gin.Context) {
	id := c.Param("id")
	var requestBody arsipRequest
	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request data"})
		return
	}

	var arsip models.Arsip
	if err := initializers.DB.First(&arsip, id).Error; err != nil {
		c.JSON(http.StatusNotFound, gin.H{"error": "Arsip not found"})
		return
	}

	if requestBody.TanggalDokumen != nil {
		tanggal, err := time.Parse("2006-01-02", *requestBody.TanggalDokumen)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		arsip.TanggalDokumen = &tanggal
	}

	// Update fields if provided in request
	if requestBody.NoArsip != nil {
		arsip.NoArsip = requestBody.NoArsip
	}
	if requestBody.JenisDokumen != nil {
		arsip.JenisDokumen = requestBody.JenisDokumen
	}
	if requestBody.NoDokumen != nil {
		arsip.NoDokumen = requestBody.NoDokumen
	}
	if requestBody.Perihal != nil {
		arsip.Perihal = requestBody.Perihal
	}
	if requestBody.NoBox != nil {
		arsip.NoBox = requestBody.NoBox
	}
	if requestBody.Keterangan != nil {
		arsip.Keterangan = requestBody.Keterangan
	}
	if requestBody.CreateBy != "" {
		arsip.CreateBy = requestBody.CreateBy
	}

	if err := initializers.DB.Save(&arsip).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to update arsip"})
		return
	}
	c.JSON(http.StatusOK, gin.H{"arsip": arsip})
}

func ArsipDelete(c *gin.Context) {
	id := c.Param("id")
	var arsip models.Arsip
	if err := initializers.DB.Where("id = ?", id).Delete(&arsip).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to delete arsip"})
		return
	}
	c.JSON(http.StatusOK, gin.H{"message": "Arsip deleted successfully"})
}

func CreateExcelArsip(c *gin.Context) {
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
	sheetNames := []string{"MEMO", "BERITA ACARA", "SK", "SURAT", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}

	// Create sheets and set headers for "ARSIP" only
	for _, sheetName := range sheetNames {
		if sheetName == "ARSIP" {
			f.NewSheet(sheetName)
			f.SetCellValue(sheetName, "A1", "No Arsip")
			f.SetCellValue(sheetName, "B1", "Jenis Dokumen")
			f.SetCellValue(sheetName, "C1", "No Dokumen")
			f.SetCellValue(sheetName, "D1", "Perihal")
			f.SetCellValue(sheetName, "E1", "No Box")
			f.SetCellValue(sheetName, "F1", "Keterangan")
			f.SetCellValue(sheetName, "G1", "Tanggal Dokumen")
			f.SetCellValue(sheetName, "H1", "Tanggal Penyerahan")

			// Set column widths for better readability
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

	err = f.SetCellStyle("ARSIP", "A1", "H1", styleHeader)

	// Fetch initial data from the database
	var arsips []models.Arsip
	initializers.DB.Find(&arsips)

	// Write initial data to the "ARSIP" sheet
	// ... existing code ...

	// Write initial data to the "ARSIP" sheet
	arsipSheetName := "ARSIP"
	for i, arsip := range arsips {
		rowNum := i + 2 // Start from the second row (first row is header)
		f.SetCellValue(arsipSheetName, fmt.Sprintf("A%d", rowNum), derefString(arsip.NoArsip))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("B%d", rowNum), derefString(arsip.JenisDokumen))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("C%d", rowNum), derefString(arsip.NoDokumen))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("D%d", rowNum), derefString(arsip.Perihal))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("E%d", rowNum), derefString(arsip.NoBox))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("F%d", rowNum), derefString(arsip.Keterangan))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("G%d", rowNum), arsip.TanggalDokumen.Format("2006-01-02"))
		f.SetCellValue(arsipSheetName, fmt.Sprintf("H%d", rowNum), arsip.TanggalPenyerahan.Format("2006-01-02"))
	}

	styleAll, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return
	}

	err = f.SetCellStyle("ARSIP", "A2", "H"+strconv.Itoa(len(arsips)+1), styleAll)

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

func derefString(s *string) string {
	if s != nil {
		return *s
	}
	return ""
}

func UpdateSheetArsip(c *gin.Context) {
	dir := "C:\\excel"
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

	sheetName := "ARSIP"

	if _, err := f.GetSheetIndex(sheetName); err == nil {
		f.DeleteSheet(sheetName)
	}
	f.NewSheet(sheetName)

	f.SetCellValue(sheetName, "A1", "No Arsip")
	f.SetCellValue(sheetName, "B1", "Jenis Dokumen")
	f.SetCellValue(sheetName, "C1", "No Dokumen")
	f.SetCellValue(sheetName, "D1", "Perihal")
	f.SetCellValue(sheetName, "E1", "No Box")
	f.SetCellValue(sheetName, "F1", "Keterangan")
	f.SetCellValue(sheetName, "G1", "Tanggal Dokumen")
	f.SetCellValue(sheetName, "H1", "Tanggal Penyerahan")

	var arsips []models.Arsip
	initializers.DB.Find(&arsips)

	for i, arsip := range arsips {
		rowNum := i + 2
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), arsip.NoArsip)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), arsip.JenisDokumen)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), arsip.NoDokumen)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), arsip.Perihal)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), arsip.NoBox)
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), arsip.Keterangan)
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), arsip.TanggalDokumen.Format("2006-01-02"))
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), arsip.TanggalPenyerahan.Format("2006-01-02"))
	}

	if err := f.SaveAs(filePath); err != nil {
		c.String(http.StatusInternalServerError, "Error menyimpan file: %v", err)
		return
	}
}

func ImportExcelArsip(c *gin.Context) {
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
	sheetName := "ARSIP"
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

		noArsip := row[0]
		jenisDokumen := row[1]
		noDokumen := row[2]
		perihal := row[3]
		noBox := row[4]
		keterangan := row[5]
		tanggalDokumen := row[6]
		tanggalPenyerahan := row[7]

		// Parse tanggal
		tanggalDokumenString, err := time.Parse("2006-01-02", tanggalDokumen)
		if err != nil {
			c.String(http.StatusBadRequest, "Invalid date format in row %d: %v", i+1, err)
			return
		}
		tanggalPenyerahanString, err := time.Parse("2006-01-02", tanggalPenyerahan)
		if err != nil {
			c.String(http.StatusBadRequest, "Invalid date format in row %d: %v", i+1, err)
			return
		}

		arsip := models.Arsip{
			NoArsip:           &noArsip,
			JenisDokumen:      &jenisDokumen,
			NoDokumen:         &noDokumen,
			Perihal:           &perihal,
			NoBox:             &noBox,
			Keterangan:        &keterangan,
			TanggalDokumen:    &tanggalDokumenString,
			TanggalPenyerahan: &tanggalPenyerahanString,
			CreateBy:          c.MustGet("username").(string),
		}

		// Simpan ke database
		if err := initializers.DB.Create(&arsip).Error; err != nil {
			log.Printf("Error saving record from row %d: %v", i+1, err)
			c.String(http.StatusInternalServerError, "Error saving record from row %d: %v", i+1, err)
			return
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully."})
}
