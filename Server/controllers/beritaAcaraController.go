package controllers

import (
	"context"
	"errors"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"project-its/initializers"
	"project-its/models"
	"strconv"
	"strings"
	"time"

	"github.com/Azure/azure-storage-blob-go/azblob"
	"github.com/gin-gonic/gin"
	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
	"gorm.io/gorm"
)

type BcRequest struct {
	ID       uint    `gorm:"primaryKey"`
	Tanggal  *string `json:"tanggal"`
	NoSurat  string  `json:"no_surat"`
	Perihal  *string `json:"perihal"`
	Pic      *string `json:"pic"`
	CreateBy string  `json:"create_by"`
}

func init() {
	err := godotenv.Load() // Memuat file .env
	if err != nil {
		log.Fatal("Error loading .env file")
	}

	accountNameBeritaAcara = os.Getenv("ACCOUNT_NAME") // Mengambil nilai dari .env
	accountKeyBeritaAcara = os.Getenv("ACCOUNT_KEY")   // Mengambil nilai dari .env
	containerNameBeritaAcara = "beritaacaraits"        // Mengambil nilai dari .env
}

// Tambahkan variabel global untuk menyimpan kredensial
var (
	accountNameBeritaAcara   string
	accountKeyBeritaAcara    string
	containerNameBeritaAcara string
)

func getBlobServiceClientBeritaAcara() azblob.ServiceURL {
	creds, err := azblob.NewSharedKeyCredential(accountNameBeritaAcara, accountKeyBeritaAcara)
	if err != nil {
		panic("Failed to create shared key credential: " + err.Error())
	}

	pipeline := azblob.NewPipeline(creds, azblob.PipelineOptions{})

	// Build the URL for the Azure Blob Storage account
	URL, err := url.Parse(fmt.Sprintf("https://%s.blob.core.windows.net/", accountNameBeritaAcara))
	if err != nil {
		log.Fatal("Invalid URL format")
	}

	// Create a ServiceURL object that wraps the URL and the pipeline
	serviceURL := azblob.NewServiceURL(*URL, pipeline)

	return serviceURL
}

func UploadHandlerBeritaAcara(c *gin.Context) {
	id := c.PostForm("id") // Mendapatkan ID dari form data
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "File diperlukan"})
		return
	}

	// Membuat path berdasarkan ID
	filename := fmt.Sprintf("%s/%s", id, file.Filename)

	// Membuka file
	src, err := file.Open()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal membuka file"})
		return
	}
	defer src.Close()

	// Mengunggah file ke Azure Blob Storage
	containerURL := getBlobServiceClientBeritaAcara().NewContainerURL(containerNameBeritaAcara)
	blobURL := containerURL.NewBlockBlobURL(filename)

	_, err = azblob.UploadStreamToBlockBlob(context.TODO(), src, blobURL, azblob.UploadStreamToBlockBlobOptions{})
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal mengunggah file"})
		return
	}

	// Menambahkan log untuk menunjukkan ke kontainer mana file diunggah
	log.Printf("File %s berhasil diunggah ke kontainer %s", filename, containerNameBeritaAcara)

	c.JSON(http.StatusOK, gin.H{"message": "File berhasil diunggah"})
}

func GetFilesByIDBeritaAcara(c *gin.Context) {
	id := c.Param("id") // Mendapatkan ID dari URL

	containerURL := getBlobServiceClient().NewContainerURL(containerNameBeritaAcara)
	prefix := fmt.Sprintf("%s/", id) // Prefix untuk daftar blob di folder tertentu (ID)

	var files []string
	for marker := (azblob.Marker{}); marker.NotDone(); {
		listBlob, err := containerURL.ListBlobsFlatSegment(context.TODO(), marker, azblob.ListBlobsSegmentOptions{
			Prefix: prefix, // Hanya daftar blob dengan prefix yang ditentukan (folder)
		})
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "Gagal membuat daftar file"})
			return
		}

		for _, blobInfo := range listBlob.Segment.BlobItems {
			files = append(files, blobInfo.Name)
		}

		marker = listBlob.NextMarker
	}

	c.JSON(http.StatusOK, gin.H{"files": files}) // Pastikan mengembalikan array files
}

// Fungsi untuk menghapus file dari Azure Blob Storage
func DeleteFileHandlerBeritaAcara(c *gin.Context) {
	filename := c.Param("filename")
	id := c.Param("id")
	if filename == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Filename is required"})
		return
	}

	// Membuat path lengkap berdasarkan ID dan nama file
	fullPath := fmt.Sprintf("%s/%s", id, filename)

	containerURL := getBlobServiceClient().NewContainerURL(containerNameBeritaAcara)
	blobURL := containerURL.NewBlockBlobURL(fullPath)

	// Menghapus blob
	_, err := blobURL.Delete(context.TODO(), azblob.DeleteSnapshotsOptionNone, azblob.BlobAccessConditions{})
	if err != nil {
		log.Printf("Error deleting file: %v", err) // Log kesalahan
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to delete file"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "File deleted successfully"}) // Pastikan ini ada
}

// Fungsi untuk mendownload file dari Azure Blob Storage
func DownloadFileHandlerBeritaAcara(c *gin.Context) {
	id := c.Param("id") // Mendapatkan ID dari URL
	filename := c.Param("filename")
	if filename == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Filename is required"})
		return
	}

	// Membuat path lengkap berdasarkan ID dan nama file
	fullPath := fmt.Sprintf("%s/%s", id, filename)

	containerURL := getBlobServiceClient().NewContainerURL(containerNameMemo)
	blobURL := containerURL.NewBlockBlobURL(fullPath)

	downloadResponse, err := blobURL.Download(context.TODO(), 0, azblob.CountToEnd, azblob.BlobAccessConditions{}, false, azblob.ClientProvidedKeyOptions{})
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to download file"})
		return
	}

	bodyStream := downloadResponse.Body(azblob.RetryReaderOptions{})
	defer bodyStream.Close()

	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))
	c.Header("Content-Type", "application/octet-stream")

	// Mengirimkan data file ke client
	io.Copy(c.Writer, bodyStream)
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

	nomor, err := GetLatestBeritaAcaraNumber(requestBody.NoSurat)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest memo number"})
		return
	}

	// Langsung gunakan `nomor` yang sudah diformat dengan benar
	requestBody.NoSurat = nomor
	log.Printf("Generated NoMemo: %s", requestBody.NoSurat) // Log nomor memo

	requestBody.CreateBy = c.MustGet("username").(string)

	bc := models.BeritaAcara{
		Tanggal:  tanggal,
		NoSurat:  &requestBody.NoSurat,
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

	nomor, err := GetLatestMemoNumber(requestBody.NoSurat)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to get latest Berita Acara number"})
		return
	}

	// Cek apakah nomor yang diterima adalah "00001"
	if nomor == "00001" {
		// Jika "00001", berarti ini adalah entri pertama
		log.Println("This is the first Berita Acara entry.")
	}

	tahun := time.Now().Year()

	// Menentukan format NoMemo berdasarkan kategori
	if requestBody.NoSurat == "ITS-SAG" {
		noSurat := fmt.Sprintf("%s-ITS-SAG-M-%d", nomor, tahun)
		requestBody.NoSurat = noSurat
		log.Printf("Generated NoSurat for ITS-SAG: %s", requestBody.NoSurat) // Log nomor memo
	} else if requestBody.NoSurat == "ITS-ISO" {
		noSurat := fmt.Sprintf("%s-ITS-ISO-M-%d", nomor, tahun)
		requestBody.NoSurat = noSurat
		log.Printf("Generated NoSurat for ITS-ISO: %s", requestBody.NoSurat) // Log nomor memo
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

	if requestBody.NoSurat != "" {
		bc.NoSurat = &requestBody.NoSurat
	} else {
		bc.NoSurat = bc.NoSurat
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

			beritaAcaraSAG := models.BeritaAcara{
				Tanggal:  &tanggalSAG,
				NoSurat:   &noSuratSAG,
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

			beritaAcaraISO := models.BeritaAcara{
				Tanggal:  &tanggalISO,
				NoSurat:   &noSuratISO,
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
	}

	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}
