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

type BcRequest struct {
	ID       uint    `gorm:"primaryKey"`
	Tanggal  *string `json:"tanggal"`
	NoSurat  *string  `json:"no_surat"`
	Perihal  *string `json:"perihal"`
	Pic      *string `json:"pic"`
	CreateBy string  `json:"create_by"`
}

func UploadHandlerBeritaAcara(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/beritaacara"
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

func GetFilesByIDBeritaAcara(c *gin.Context) {
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

func DeleteFileHandlerBeritaAcara(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/beritaacara"
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

func DownloadFileHandlerBeritaAcara(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/beritaacara"
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

func BeritaAcaraIndex(c *gin.Context) {

	var beritaAcaras []models.BeritaAcara
	if err := initializers.DB.Find(&beritaAcaras).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal mengambil data berita acara: " + err.Error()})
		return
	}
	c.JSON(http.StatusOK, gin.H{"beritaAcaras": beritaAcaras})
}
func GetLatestBeritaAcaraNumber(category string) (string, error) {
	var lastBeritaAcara models.BeritaAcara
	currentYear := time.Now().Year()
	searchPattern := fmt.Sprintf("%%/%s/BA/%d", category, currentYear) // Mencari format seperti '%/ITS-SAG/BA/2024'

	if err := initializers.DB.Where("no_surat LIKE ?", searchPattern).Order("id desc").First(&lastBeritaAcara).Error; err != nil {
		if errors.Is(err, gorm.ErrRecordNotFound) {
			return fmt.Sprintf("00001/%s/BA/%d", category, currentYear), nil // Jika tidak ada catatan, kembalikan nomor pertama
		}
		return "", err
	}

	parts := strings.Split(*lastBeritaAcara.NoSurat, "/")
	if len(parts) > 0 {
		number, err := strconv.Atoi(parts[0])
		if err != nil {
			return "", err
		}
		return fmt.Sprintf("%05d/%s/BA/%d", number+1, category, currentYear), nil // Tambahkan 1 ke nomor terakhir
	}

	return fmt.Sprintf("00001/%s/BA/%d", category, currentYear), nil
}

func BeritaAcaraCreate(c *gin.Context) {
	var requestBody BcRequest

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

	log.Printf("Parsed date: %v", tanggal)

	nomor, err := GetLatestBeritaAcaraNumber(*requestBody.NoSurat)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest memo number"})
		return
	}

	// Langsung gunakan `nomor` yang sudah diformat dengan benar
	requestBody.NoSurat = &nomor
	log.Printf("Generated NoMemo: %s", requestBody.NoSurat) // Log nomor memo

	requestBody.CreateBy = c.MustGet("username").(string)

	bc := models.BeritaAcara{
		Tanggal:  tanggal,
		NoSurat:  requestBody.NoSurat,
		Perihal:  requestBody.Perihal,
		Pic:      requestBody.Pic,
		CreateBy: requestBody.CreateBy,
	}

	result := initializers.DB.Create(&bc)
	if result.Error != nil {
		log.Printf("Error saving memo: %v", result.Error)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create Memo Sag"})
		return
	}
	log.Printf("Memo created successfully: %v", bc)

	c.JSON(201, gin.H{
		"beritaAcara": bc,
	})
}

func BeritaAcaraShow(c *gin.Context) {

	id := c.Params.ByName("id")

	var bc models.BeritaAcara

	initializers.DB.First(&bc, id)

	c.JSON(200, gin.H{
		"beritaAcara": bc,
	})

}

func BeritaAcaraUpdate(c *gin.Context) {

	var requestBody BcRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(400, gin.H{"error": "Invalid request body: " + err.Error()})
		return
	}

	id := c.Params.ByName("id")

	var bc models.BeritaAcara

	if err := initializers.DB.First(&bc, id).Error; err != nil {
		c.JSON(404, gin.H{"error": "Berita Acara not found"})
		return
	}

	nomor, err := GetLatestBeritaAcaraNumber(*requestBody.NoSurat)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest Berita Acara number"})
		return
	}


	requestBody.CreateBy = c.MustGet("username").(string)
	bc.CreateBy = requestBody.CreateBy

	if requestBody.Tanggal != nil {
		tanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(400, gin.H{"error": "Format tanggal tidak valid: " + err.Error()})
			return
		}
		bc.Tanggal = &tanggal
	}

	if requestBody.NoSurat != nil && *requestBody.NoSurat != "" {
		bc.NoSurat = &nomor
	}

	if requestBody.Perihal != nil {
		bc.Perihal = requestBody.Perihal
	} else {
		bc.Perihal = bc.Perihal
	}

	if requestBody.Pic != nil {
		bc.Pic = requestBody.Pic
	} else {
		bc.Pic = bc.Pic
	}

	if requestBody.CreateBy != "" {
		bc.CreateBy = requestBody.CreateBy
	} else {
		bc.CreateBy = bc.CreateBy
	}

	initializers.DB.Save(&bc)

	c.JSON(200, gin.H{
		"beritaAcara": bc,
	})
}

func BeritaAcaraDelete(c *gin.Context) {

	id := c.Params.ByName("id")

	var bc models.BeritaAcara

	if err := initializers.DB.First(&bc, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Berita Acara not found"})
		return
	}

	if err := initializers.DB.Delete(&bc).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete BeritaAcara: " + err.Error()})
		return
	}

	c.Status(204)

}

func exportBeritaAcaraToExcel(beritaAcaras []models.BeritaAcara) (*excelize.File, error) {
	// Buat file Excel baru
	f := excelize.NewFile()

	sheetNames := []string{"MEMO", "BERITA ACARA", "SK", "SURAT", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}

	for _, sheetName := range sheetNames {
		f.NewSheet(sheetName)
		if sheetName == "BERITA ACARA" {
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
		}
	}
	f.DeleteSheet("Sheet1")

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

	return f, nil
}

// Handler untuk melakukan export Excel dengan Gin
func ExportBeritaAcaraHandler(c *gin.Context) {
	// Data memo contoh
	var beritaAcaras []models.BeritaAcara
	initializers.DB.Find(&beritaAcaras)

	// Buat file Excel
	f, err := exportBeritaAcaraToExcel(beritaAcaras)
	if err != nil {
		c.String(http.StatusInternalServerError, "Gagal mengekspor data ke Excel")
		return
	}

	// Set nama file dan header untuk download
	fileName := fmt.Sprintf("its_report.xlsx")
	c.Header("Content-Disposition", "attachment; filename="+fileName)
	c.Header("Content-Type", "application/octet-stream")

	// Simpan file Excel ke dalam buffer
	if err := f.Write(c.Writer); err != nil {
		c.String(http.StatusInternalServerError, "Gagal menyimpan file Excel")
	}
}

func ImportExcelBeritaAcara(c *gin.Context) {
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

	sheetName := "BERITA ACARA"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Println("Processing rows...")

	// Definisikan semua format tanggal yang mungkin
	dateFormats := []string{
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
		"06-Jan-02",
	}

	for i, row := range rows {
		if i == 0 { // Lewati baris pertama yang merupakan header
			continue
		}
		if len(row) < 4 { // Pastikan ada cukup kolom
			log.Printf("Row %d skipped: less than 4 columns filled", i+1)
			continue
		}

		// Ambil data dari kolom SAG (kiri)
		tanggalSAGStr := row[0]
		noSuratSAG := row[1]
		perihalSAG := row[2]
		picSAG := row[3]

		var tanggalSAG time.Time
		var parseErr error

		// Coba konversi dari serial Excel jika tanggalSAGStr adalah angka
		if serial, err := strconv.Atoi(tanggalSAGStr); err == nil {
			tanggalSAG, parseErr = excelDateToTimeMemo(serial)
		} else {
			// Coba parse menggunakan format tanggal yang sudah ada
			for _, format := range dateFormats {
				tanggalSAG, parseErr = time.Parse(format, tanggalSAGStr)
				if parseErr == nil {
					break // Keluar dari loop jika parsing berhasil
				}
			}
		}

		if parseErr != nil {
			log.Printf("Format tanggal tidak valid di baris %d: %v", i+1, parseErr)
			continue // Lewati baris ini jika format tanggal tidak valid
		}

		beritaAcaraSAG := models.BeritaAcara{
			Tanggal:  &tanggalSAG,
			NoSurat:  &noSuratSAG,
			Perihal:  &perihalSAG,
			Pic:      &picSAG,
			CreateBy: c.MustGet("username").(string),
		}

		if err := initializers.DB.Create(&beritaAcaraSAG).Error; err != nil {
			log.Printf("Error saving SAG record from row %d: %v", i+1, err)
		} else {
			log.Printf("SAG Row %d imported successfully", i+1)
		}
	}

	// Proses data ISO
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) < 8 { // Pastikan ada cukup kolom untuk ISO
			log.Printf("Row %d skipped: less than 8 columns filled", i+1)
			continue
		}

		// Ambil data dari kolom ISO (kanan)
		tanggalISOStr := row[5]
		noSuratISO := row[6]
		perihalISO := row[7]
		picISO := row[8]

		var tanggalISO time.Time
		var parseErr error

		// Coba konversi dari serial Excel jika tanggalISOStr adalah angka
		if serial, err := strconv.Atoi(tanggalISOStr); err == nil {
			tanggalISO, parseErr = excelDateToTimeMemo(serial)
		} else {
			// Coba parse menggunakan format tanggal yang sudah ada
			for _, format := range dateFormats {
				tanggalISO, parseErr = time.Parse(format, tanggalISOStr)
				if parseErr == nil {
					break // Keluar dari loop jika parsing berhasil
				}
			}
		}

		if parseErr != nil {
			log.Printf("Format tanggal tidak valid di baris %d: %v", i+1, parseErr)
			continue // Lewati baris ini jika format tanggal tidak valid
		}

		beritaAcaraISO := models.BeritaAcara{
			Tanggal:  &tanggalISO,
			NoSurat:  &noSuratISO,
			Perihal:  &perihalISO,
			Pic:      &picISO,
			CreateBy: c.MustGet("username").(string),
		}

		if err := initializers.DB.Create(&beritaAcaraISO).Error; err != nil {
			log.Printf("Error saving ISO record from row %d: %v", i+1, err)
		} else {
			log.Printf("ISO Row %d imported successfully", i+1)
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}