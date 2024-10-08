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

type SKRequest struct {
	ID       uint    `gorm:"primaryKey"`
	Tanggal  *string `json:"tanggal"`
	NoSurat  *string `json:"no_surat"`
	Perihal  *string `json:"perihal"`
	Pic      *string `json:"pic"`
	CreateBy string  `json:"create_by"`
}

func UploadHandlerSk(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/sk"
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

func GetFilesByIDSk(c *gin.Context) {
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

func DeleteFileHandlerSk(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/sk"
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

func DownloadFileHandlerSk(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/sk"
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

func GetLatestSuratSkNumber(NoSurat string) (string, error) {
	var lastSurat models.Sk
	// Ubah pencarian untuk menggunakan format yang benar
	searchPattern := fmt.Sprintf("%%/%s/SK/%%", NoSurat) // Ini akan mencari format seperti '%ITS-SAG/Sk/%'
	if err := initializers.DB.Where("no_surat LIKE ?", searchPattern).Order("id desc").First(&lastSurat).Error; err != nil {
		if errors.Is(err, gorm.ErrRecordNotFound) {
			return "00001", nil // Jika tidak ada catatan, kembalikan 00001
		}
		return "", err
	}

	// Ambil nomor surat terakhir, pisahkan, dan tambahkan 1
	parts := strings.Split(*lastSurat.NoSurat, "/")
	if len(parts) > 0 {
		// Ambil bagian pertama dari parts yang seharusnya adalah nomor
		numberPart := parts[0]
		number, err := strconv.Atoi(numberPart)
		if err != nil {
			log.Printf("Error parsing number from surat: %v", err)
			return "", err
		}
		return fmt.Sprintf("%05d", number+1), nil // Tambahkan 1 ke nomor terakhir
	}

	return "00001", nil
}

func SkIndex(c *gin.Context) {

	var sK []models.Sk

	initializers.DB.Find(&sK)

	c.JSON(200, gin.H{
		"sk": sK,
	})

}

func SkCreate(c *gin.Context) {
	var requestBody SKRequest

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

	nomor, err := GetLatestSuratSkNumber(*requestBody.NoSurat)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest surat number"})
		return
	}

	// Cek apakah nomor yang diterima adalah "00001"
	if nomor == "00001" {
		// Jika "00001", berarti ini adalah entri pertama
		log.Println("This is the first memo entry.")
	}

	tahun := time.Now().Year()

	// Menentukan format NoMemo berdasarkan kategori
	if *requestBody.NoSurat == "ITS-SAG" {
		noSurat := fmt.Sprintf("%s/ITS-SAG/SK/%d", nomor, tahun)
		requestBody.NoSurat = &noSurat
		log.Printf("Generated NoMemo for ITS-SAG: %s", *requestBody.NoSurat) // Log nomor surat
	} else if *requestBody.NoSurat == "ITS-ISO" {
		noSurat := fmt.Sprintf("%s/ITS-ISO/SK/%d", nomor, tahun)
		requestBody.NoSurat = &noSurat
		log.Printf("Generated NoMemo for ITS-ISO: %s", *requestBody.NoSurat) // Log nomor surat
	}

	requestBody.CreateBy = c.MustGet("username").(string)

	sK := models.Sk{
		Tanggal:  tanggal,             // Gunakan tanggal yang telah diparsing, bisa jadi nil jika input kosong
		NoSurat:  requestBody.NoSurat, // Menggunakan NoMemo yang sudah diformat
		Perihal:  requestBody.Perihal,
		Pic:      requestBody.Pic,
		CreateBy: requestBody.CreateBy,
	}

	result := initializers.DB.Create(&sK)
	if result.Error != nil {
		log.Printf("Error saving surat: %v", result.Error)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create Memo Sag"})
		return
	}
	log.Printf("Memo created successfully: %v", sK)

	c.JSON(201, gin.H{
		"sk": sK,
	})
}

func SkShow(c *gin.Context) {

	id := c.Params.ByName("id")

	var sK models.Sk

	initializers.DB.First(&sK, id)

	c.JSON(200, gin.H{
		"sk": sK,
	})

}

func SkUpdate(c *gin.Context) {
	var requestBody SuratRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request body: " + err.Error()})
		return
	}

	id := c.Param("id")
	var surat models.Surat

	// Cari memo berdasarkan ID
	if err := initializers.DB.First(&surat, id).Error; err != nil {
		log.Printf("Memo with ID %s not found: %v", id, err)
		c.JSON(http.StatusNotFound, gin.H{"error": "Memo not found"})
		return
	}

	nomor, err := GetLatestMemoNumber(*requestBody.NoSurat)
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
	// Menentukan format NoSurat berdasarkan kategori
	if *requestBody.NoSurat == "ITS-SAG" {
		noSurat := fmt.Sprintf("%s/ITS-SAG/M/%d", nomor, tahun)
		requestBody.NoSurat = &noSurat
		log.Printf("Generated NoSurat for ITS-SAG: %s", *requestBody.NoSurat) // Log nomor Surat
	} else if *requestBody.NoSurat == "ITS-ISO" {
		noSurat := fmt.Sprintf("%s/ITS-ISO/M/%d", nomor, tahun)
		requestBody.NoSurat = &noSurat
		log.Printf("Generated NoSurat for ITS-ISO: %s", *requestBody.NoSurat) // Log nomor memo
	}

	// Update tanggal jika diberikan dan tidak kosong
	if requestBody.Tanggal != nil && *requestBody.Tanggal != "" {
		parsedTanggal, err := time.Parse("2006-01-02", *requestBody.Tanggal)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		surat.Tanggal = &parsedTanggal
	}

	// Update nomor jika diberikan dan tidak kosong
	if requestBody.NoSurat != nil && *requestBody.NoSurat != "" {
		surat.NoSurat = requestBody.NoSurat
	}

	// Update fields lainnya
	if requestBody.Perihal != nil {
		surat.Perihal = requestBody.Perihal
	}
	if requestBody.Pic != nil {
		surat.Pic = requestBody.Pic
	}

	// Simpan perubahan
	if err := initializers.DB.Save(&surat).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to update surat"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "surat updated successfully", "surat": surat})
}

func SkDelete(c *gin.Context) {

	id := c.Params.ByName("id")

	var sK models.Sk

	if err := initializers.DB.First(&sK, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Memo not found"})
		return
	}

	if err := initializers.DB.Delete(&sK).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete Memo: " + err.Error()})
		return
	}

	c.Status(204)

}

func exportSkToExcel(sKs []models.Sk) (*excelize.File, error) {
	// Buat file Excel baru
	f := excelize.NewFile()

	sheetNames := []string{"MEMO", "BERITA ACARA", "SK", "SURAT", "PROJECT", "PERDIN", "SURAT MASUK", "SURAT KELUAR", "ARSIP", "MEETING", "MEETING SCHEDULE"}

	for _, sheetName := range sheetNames {
		f.NewSheet(sheetName)
		if sheetName == "SK" {
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
	for _, sK := range sKs {
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

		// Pisahkan NoSurat untuk mendapatkan tipe surat
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

	return f, nil
}

// Handler untuk melakukan export Excel dengan Gin
func ExportSkHandler(c *gin.Context) {
	// Data memo contoh
	var sKs []models.Sk
	initializers.DB.Find(&sKs)

	// Buat file Excel
	f, err := exportSkToExcel(sKs)
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

func ImportExcelSk(c *gin.Context) {
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

	sheetName := "SK"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Println("Processing rows...")
	for i := range rows {
		if i == 0 {
			continue
		}

		// Ambil data dari kolom SAG (kiri)
		var (
			tanggalSAGStr, _ = f.GetCellValue(sheetName, fmt.Sprintf("A%d", i+1))
			noSuratSAG, _     = f.GetCellValue(sheetName, fmt.Sprintf("B%d", i+1))
			perihalSAG, _    = f.GetCellValue(sheetName, fmt.Sprintf("C%d", i+1))
			picSAG, _        = f.GetCellValue(sheetName, fmt.Sprintf("D%d", i+1))
		)

		// Proses data SAG jika ada
		if tanggalSAGStr != "" || noSuratSAG != "" || perihalSAG != "" || picSAG != "" {
			tanggalSAG, err := parseDate(tanggalSAGStr)
			if err != nil {
				log.Printf("Error parsing SAG date from row %d: %v", i+1, err)

			}

			skSAG := models.Sk{
				Tanggal:  &tanggalSAG,
				NoSurat:   &noSuratSAG,
				Perihal:  &perihalSAG,
				Pic:      &picSAG,
				CreateBy: c.MustGet("username").(string),
			}

			if err := initializers.DB.Create(&skSAG).Error; err != nil {
				log.Printf("Error saving SAG record from row %d: %v", i+1, err)

			} else {
				log.Printf("SAG Row %d imported successfully", i+1)
			}
		}
	}
	// Proses data ISO
	isoRows, err := f.GetRows(sheetName)
	if err != nil {
		c.String(http.StatusInternalServerError, "Error getting ISO rows: %v", err)
		return
	}
	for i := range isoRows {
		if i == 0 {
			continue
		}

		// Ambil data dari kolom ISO (kanan)
		var (
			tanggalISOStr, _ = f.GetCellValue(sheetName, fmt.Sprintf("F%d", i+1))
			noSuratISO, _     = f.GetCellValue(sheetName, fmt.Sprintf("G%d", i+1))
			perihalISO, _    = f.GetCellValue(sheetName, fmt.Sprintf("H%d", i+1))
			picISO, _        = f.GetCellValue(sheetName, fmt.Sprintf("I%d", i+1))
		)

		// Proses data ISO jika ada
		if tanggalISOStr != "" || noSuratISO != "" || perihalISO != "" || picISO != "" {
			tanggalISO, err := parseDate(tanggalISOStr)
			if err != nil {
				log.Printf("Error parsing ISO date from row %d: %v", i+1, err)

			}

			skISO := models.Sk{
				Tanggal:  &tanggalISO,
				NoSurat:   &noSuratISO,
				Perihal:  &perihalISO,
				Pic:      &picISO,
				CreateBy: c.MustGet("username").(string),
			}

			if err := initializers.DB.Create(&skISO).Error; err != nil {
				log.Printf("Error saving ISO record from row %d: %v", i+1, err)

			} else {
				log.Printf("ISO Row %d imported successfully", i+1)
			}
		}
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}
