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
	"strings"
	"time"
	"unicode"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type ProjectRequest struct {
	ID              uint    `gorm:"primaryKey"`
	KodeProject     *string `json:"kode_project"`
	JenisPengadaan  *string `json:"jenis_pengadaan"`
	NamaPengadaan   *string `json:"nama_pengadaan"`
	DivInisiasi     *string `json:"div_inisiasi"`
	Bulan           *string `json:"bulan"`
	SumberPendanaan *string `json:"sumber_pendanaan"`
	Anggaran        *string `json:"anggaran"`
	NoIzin          *string `json:"no_izin"`
	TanggalIzin     *string `json:"tanggal_izin"`
	TanggalTor      *string `json:"tanggal_tor"`
	Pic             *string `json:"pic"`
	CreateBy        string  `json:"create_by"`
	Group           *string `json:"group"`
	InfraType       *string `json:"infra_type"`
	BudgetType      *string `json:"budget_type"`
	Type            *string `json:"type"`
}

func UploadHandlerProject(c *gin.Context) {
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

	baseDir := "C:/UploadedFile/project"
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

func GetFilesByIDProject(c *gin.Context) {
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

func DeleteFileHandlerProject(c *gin.Context) {
	encodedFilename := c.Param("filename")
	filename, err := url.QueryUnescape(encodedFilename)
	if err != nil {
		log.Printf("Error decoding filename: %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid filename"})
		return
	}

	id := c.Param("id")
	log.Printf("Received ID: %s and Filename: %s", id, filename) // Tambahkan log ini

	baseDir := "C:/UploadedFile/project"
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

func DownloadFileHandlerProject(c *gin.Context) {
	id := c.Param("id")
	filename := c.Param("filename")
	baseDir := "C:/UploadedFile/project"
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

func ProjectCreate(c *gin.Context) {
	var requestBody ProjectRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request"})
		return
	}

	// Generate KodeProject based on Group
	var lastNumber int
	var newKodeProject string
	currentYear := time.Now().Format("2006")

	if requestBody.Group != nil {
		// Fetch the last project of the same group and year
		lastProject := models.Project{}
		initializers.DB.Where("kode_project LIKE ?", fmt.Sprintf("%%/%s/%%/%s", *requestBody.Group, currentYear)).Order("id desc").First(&lastProject)

		if lastProject.KodeProject != nil {
			fmt.Sscanf(*lastProject.KodeProject, "%d/", &lastNumber)
		}

		newNumber := fmt.Sprintf("%05d", lastNumber+1) // Increment and format
		newKodeProject = fmt.Sprintf("%s/%s/%s/%s/%s/%s", newNumber, *requestBody.Group, *requestBody.InfraType, *requestBody.BudgetType, *requestBody.Type, currentYear)
	} else {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Group is required"})
		return
	}

	var bulan *time.Time
	if requestBody.Bulan != nil && *requestBody.Bulan != "" {
		parsedBulan, err := time.Parse("2006-01", *requestBody.Bulan)
		if err != nil {
			log.Printf("Error parsing date: %v", err)
			c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
			return
		}
		bulan = &parsedBulan
	}

	log.Printf("Parsed date: %v", bulan)

	var tanggal_izin *time.Time
	if requestBody.TanggalIzin != nil && *requestBody.TanggalIzin != "" {
		parsedTanggalIzin, err := time.Parse("2006-01-02", *requestBody.TanggalIzin)
		if err != nil {
			log.Printf("Error parsing date: %v", err)
			c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
			return
		}
		tanggal_izin = &parsedTanggalIzin
	}

	log.Printf("Parsed date: %v", tanggal_izin)

	var tanggal_tor *time.Time
	if requestBody.TanggalTor != nil && *requestBody.TanggalTor != "" {
		parsedTanggalTor, err := time.Parse("2006-01-02", *requestBody.TanggalTor)
		if err != nil {
			log.Printf("Error parsing date: %v", err)
			c.JSON(400, gin.H{"error": "Invalid date format: " + err.Error()})
			return
		}
		tanggal_tor = &parsedTanggalTor
	}

	log.Printf("Parsed date: %v", tanggal_tor)

	requestBody.KodeProject = &newKodeProject

	requestBody.CreateBy = c.MustGet("username").(string)

	project := models.Project{
		KodeProject:     requestBody.KodeProject,
		JenisPengadaan:  requestBody.JenisPengadaan,
		NamaPengadaan:   requestBody.NamaPengadaan,
		DivInisiasi:     requestBody.DivInisiasi,
		Bulan:           bulan,
		SumberPendanaan: requestBody.SumberPendanaan,
		Anggaran:        requestBody.Anggaran,
		NoIzin:          requestBody.NoIzin,
		TanggalIzin:     tanggal_izin,
		TanggalTor:      tanggal_tor,
		Pic:             requestBody.Pic,
		Group:           requestBody.Group,
		InfraType:       requestBody.InfraType,
		BudgetType:      requestBody.BudgetType,
		Type:            requestBody.Type,
		CreateBy:        requestBody.CreateBy,
	}

	// Log data project yang baru dibuat
	log.Printf("Creating new project: %+v", project)

	result := initializers.DB.Create(&project)
	if result.Error != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create project"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"project": project})
}

func ProjectIndex(c *gin.Context) {

	// Get models from DB
	var project []models.Project
	initializers.DB.Find(&project)

	//Respond with them
	c.JSON(200, gin.H{
		"project": project,
	})
}

func ProjectShow(c *gin.Context) {

	//get id
	id := c.Params.ByName("id")
	// Get models from DB
	var project models.Project

	initializers.DB.First(&project, id)

	//Respond with them
	c.JSON(200, gin.H{
		"project": project,
	})
}

func ProjectUpdate(c *gin.Context) {
	var requestBody ProjectRequest

	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request"})
		return
	}

	id := c.Params.ByName("id")
	var project models.Project
	if err := initializers.DB.First(&project, id).Error; err != nil {
		c.JSON(http.StatusNotFound, gin.H{"error": "Project not found"})
		return
	}

	// Update KodeProject if Group or other relevant fields are changed
	currentYear := time.Now().Format("2006")
	var group, infraType, budgetType, projectType string

	if requestBody.Group != nil {
		group = *requestBody.Group
	} else {
		group = *project.Group
	}

	if requestBody.InfraType != nil {
		infraType = *requestBody.InfraType
	} else {
		infraType = *project.InfraType
	}

	if requestBody.BudgetType != nil {
		budgetType = *requestBody.BudgetType
	} else {
		budgetType = *project.BudgetType
	}

	if requestBody.Type != nil {
		projectType = *requestBody.Type
	} else {
		projectType = *project.Type
	}

	lastProject := models.Project{}
	initializers.DB.Where("kode_project LIKE ?", fmt.Sprintf("%%/%s/%%/%s", group, currentYear)).Order("id desc").First(&lastProject)
	var lastNumber int
	if lastProject.KodeProject != nil {
		fmt.Sscanf(*lastProject.KodeProject, "%d/", &lastNumber)
	}
	newNumber := fmt.Sprintf("%05d", lastNumber+1)
	newKodeProject := fmt.Sprintf("%s/%s/%s/%s/%s/%s", newNumber, group, infraType, budgetType, projectType, currentYear)
	project.KodeProject = &newKodeProject

	if requestBody.Bulan != nil && *requestBody.Bulan != "" {
		parsedBulan, err := time.Parse("2006-01-02", *requestBody.Bulan)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		project.Bulan = &parsedBulan
	}

	if requestBody.TanggalIzin != nil && *requestBody.TanggalIzin != "" {
		parsedTanggal_izin, err := time.Parse("2006-01-02", *requestBody.TanggalIzin)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		project.TanggalIzin = &parsedTanggal_izin
	}

	if requestBody.TanggalTor != nil && *requestBody.TanggalTor != "" {
		parsedTanggal_tor, err := time.Parse("2006-01-02", *requestBody.TanggalTor)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid date format"})
			return
		}
		project.TanggalTor = &parsedTanggal_tor
	}

	// Update other fields
	if requestBody.JenisPengadaan != nil {
		project.JenisPengadaan = requestBody.JenisPengadaan
	}
	if requestBody.NamaPengadaan != nil {
		project.NamaPengadaan = requestBody.NamaPengadaan
	}
	if requestBody.DivInisiasi != nil {
		project.DivInisiasi = requestBody.DivInisiasi
	}
	if requestBody.SumberPendanaan != nil {
		project.SumberPendanaan = requestBody.SumberPendanaan
	}
	if requestBody.Anggaran != nil {
		project.Anggaran = requestBody.Anggaran
	}
	if requestBody.NoIzin != nil {
		project.NoIzin = requestBody.NoIzin
	}
	if requestBody.Pic != nil {
		project.Pic = requestBody.Pic
	}
	project.CreateBy = c.MustGet("username").(string)

	// Save changes
	if err := initializers.DB.Save(&project).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to update project"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"project": project})
}

func ProjectDelete(c *gin.Context) {

	//get id
	id := c.Params.ByName("id")

	var project models.Project

	if err := initializers.DB.First(&project, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Project not found"})
		return
	}

	if err := initializers.DB.Delete(&project).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete project: " + err.Error()})
		return
	}

	c.Status(204)
}

func ImportExcelProject(c *gin.Context) {
	log.Println("Starting ImportExcelProject function")

	file, _, err := c.Request.FormFile("file")
	if err != nil {
		log.Printf("Error retrieving the file: %v", err)
		c.String(http.StatusBadRequest, "Error retrieving the file: %v", err)
		return
	}
	defer file.Close()

	tempFile, err := os.CreateTemp("", "*.xlsx")
	if err != nil {
		log.Printf("Error creating temporary file: %v", err)
		c.String(http.StatusInternalServerError, "Error creating temporary file: %v", err)
		return
	}
	defer os.Remove(tempFile.Name())

	if _, err := io.Copy(tempFile, file); err != nil {
		log.Printf("Error copying file: %v", err)
		c.String(http.StatusInternalServerError, "Error copying file: %v", err)
		return
	}

	tempFile.Seek(0, 0)
	f, err := excelize.OpenFile(tempFile.Name())
	if err != nil {
		log.Printf("Error opening file: %v", err)
		c.String(http.StatusInternalServerError, "Error opening file: %v", err)
		return
	}
	defer f.Close()

	sheetName := "PROJECT"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Printf("Error getting rows: %v", err)
		c.String(http.StatusInternalServerError, "Error getting rows: %v", err)
		return
	}

	log.Printf("Total rows found: %d", len(rows))

	for i, row := range rows {
		if i < 6 { // Skip header or initial rows if necessary
			log.Printf("Skipping row %d (header or initial rows)", i+1)
			continue
		}
		// Count non-empty columns
		nonEmptyCount := 0
		for _, cell := range row {
			if cell != "" {
				nonEmptyCount++
			}
		}

		// Skip rows with less than 3 non-empty columns
		if nonEmptyCount < 3 {
			log.Printf("Row %d skipped: less than 3 columns filled, only %d filled", i+1, nonEmptyCount)
			continue
		}

		// Membersihkan string anggaran dari karakter non-numerik
		rawAnggaran := getStringOrNil(getColumn(row, 7))
		var anggaranCleaned *string
		if rawAnggaran != nil {
			cleanedAnggaran := cleanNumericString(*rawAnggaran)
			anggaranCleaned = &cleanedAnggaran
		}

		project := models.Project{
			KodeProject:     getStringOrNil(getColumn(row, 1)),
			JenisPengadaan:  getStringOrNil(getColumn(row, 2)),
			NamaPengadaan:   getStringOrNil(getColumn(row, 3)),
			DivInisiasi:     getStringOrNil(getColumn(row, 4)),
			Bulan:           parseDateOrNil(getStringOrNil(getColumn(row, 5))),
			SumberPendanaan: getStringOrNil(getColumn(row, 6)),
			Anggaran:        anggaranCleaned,
			NoIzin:          getStringOrNil(getColumn(row, 8)),
			TanggalIzin:     parseDateOrNil(getStringOrNil(getColumn(row, 9))),
			TanggalTor:      parseDateOrNil(getStringOrNil(getColumn(row, 10))),
			Pic:             getStringOrNil(getColumn(row, 11)),
			CreateBy:        c.MustGet("username").(string),
		}

		// Log data yang diimpor
		log.Printf("Importing row %d", i+1)

		if err := initializers.DB.Create(&project).Error; err != nil {
			log.Printf("Error saving record from row %d: %v", i+1, err)
			continue
		} else {
			log.Printf("Record from row %d saved successfully", i+1)
		}
	}

	log.Println("ImportExcelProject function completed")
	c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully, check logs for any skipped rows."})
}

// Helper function to safely get column data or return empty if index is out of range
func getColumn(row []string, index int) string {
	if index >= len(row) {
		return ""
	}
	return row[index]
}

// Helper function to return nil if the string is empty
func getStringOrNil(value string) *string {
	if value == "" {
		return nil
	}
	return &value
}

// Helper function to parse date from various formats
func parseDate(dateStr string) (time.Time, error) {
	dateFormats := []string{
		"2 January 2006",
		"02-06",
		"2-January-2006",
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
		"01/06",
		"02/06",
		"Jan-06", // Menambahkan format ini untuk mengenali "Feb-24" sebagai "Feb-2024"
	}

	// Menambahkan logika untuk menangani format "Feb-24"
	if strings.Contains(dateStr, "-") && len(dateStr) == 5 {
		dateStr = dateStr[:3] + "20" + dateStr[4:]
	}

	for _, format := range dateFormats {
		parsedDate, err := time.Parse(format, dateStr)
		if err == nil {
			return parsedDate, nil
		}
	}
	return time.Time{}, fmt.Errorf("no valid date format found")
}

// Fungsi untuk membersihkan string dari karakter non-numerik
func cleanNumericString(input string) string {
	return strings.Map(func(r rune) rune {
		if unicode.IsDigit(r) {
			return r
		}
		return -1
	}, input)
}

// Helper function to parse date or return nil if input is nil
func parseDateOrNil(dateStr *string) *time.Time {
	if dateStr == nil {
		return nil
	}
	parsedDate, err := parseDate(*dateStr)
	if err != nil {
		return nil
	}
	return &parsedDate
}

func exportProjectToExcel(projects []models.Project) (*excelize.File, error) {
	// Create a new Excel file
	f := excelize.NewFile()

	// Create sheets
	sheetName := "PROJECT"
	f.NewSheet(sheetName)

	// Set header for SAG (left column)
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

	f.DeleteSheet("Sheet1")

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
		return nil, err
	}
	err = f.SetCellStyle("PROJECT", fmt.Sprintf("A1"), fmt.Sprintf("K1"), styleHeader)
	if err != nil {
		return nil, err
	}

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
		return nil, err
	}
	err = f.SetCellStyle("PROJECT", fmt.Sprintf("A2"), fmt.Sprintf("K2"), styleLine)
	if err != nil {
		return nil, err
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
		return nil, err
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
					return nil, err
				}
				formattedAnggaran := fmt.Sprintf("Rp. %d", anggaran)
				f.SetCellValue("PROJECT", fmt.Sprintf("G%d", rowIndex), formattedAnggaran)
				styleAnggaran, err := f.NewStyle(&excelize.Style{
					NumFmt: 3,
				})
				if err != nil {
					return nil, err
				}
				err = f.SetCellStyle("PROJECT", fmt.Sprintf("G%d", rowIndex), fmt.Sprintf("G%d", rowIndex), styleAnggaran)
				if err != nil {
					return nil, err
				}
			}

			f.SetCellValue("PROJECT", fmt.Sprintf("H%d", rowIndex), noIzin)
			f.SetCellValue("PROJECT", fmt.Sprintf("I%d", rowIndex), tanggalIzin)
			f.SetCellValue("PROJECT", fmt.Sprintf("J%d", rowIndex), tanggalTor)
			f.SetCellValue("PROJECT", fmt.Sprintf("K%d", rowIndex), pic)
			err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", rowIndex), fmt.Sprintf("K%d", rowIndex), styleAll)
			if err != nil {
				return nil, err
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
					return nil, err
				}
				formattedAnggaran := fmt.Sprintf("Rp. %d", anggaran)
				f.SetCellValue("PROJECT", fmt.Sprintf("G%d", rowISO), formattedAnggaran)
				styleAnggaran, err := f.NewStyle(&excelize.Style{
					NumFmt: 3,
				})
				if err != nil {
					return nil, err
				}
				err = f.SetCellStyle("PROJECT", fmt.Sprintf("G%d", rowISO), fmt.Sprintf("G%d", rowISO), styleAnggaran)
				if err != nil {
					return nil, err
				}
			}

			f.SetCellValue("PROJECT", fmt.Sprintf("H%d", rowISO), noIzin)
			f.SetCellValue("PROJECT", fmt.Sprintf("I%d", rowISO), tanggalIzin)
			f.SetCellValue("PROJECT", fmt.Sprintf("J%d", rowISO), tanggalTor)
			f.SetCellValue("PROJECT", fmt.Sprintf("K%d", rowISO), pic)
			err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", rowISO), fmt.Sprintf("K%d", rowISO), styleAll)
			if err != nil {
				return nil, err
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
		return nil, err
	}

	f.SetCellValue("PROJECT", fmt.Sprintf("F%d", lastRowSAG), "ISO")
	err = f.SetCellStyle("PROJECT", fmt.Sprintf("A%d", lastRowSAG), fmt.Sprintf("K%d", lastRowSAG), styleLine)
	if err != nil {
		return nil, err
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

	return f, nil
}

// Handler untuk melakukan export Excel dengan Gin
func ExportProjectHandler(c *gin.Context) {
	// Data memo contoh
	var projects []models.Project
	initializers.DB.Find(&projects)

	// Buat file Excel
	f, err := exportProjectToExcel(projects)
	if err != nil {
		c.String(http.StatusInternalServerError, "Gagal mengekspor data ke Excel")
		return
	}

	// Set nama file dan header untuk download
	fileName := fmt.Sprintf("its_report_project.xlsx")
	c.Header("Content-Disposition", "attachment; filename="+fileName)
	c.Header("Content-Type", "application/octet-stream")

	// Simpan file Excel ke dalam buffer
	if err := f.Write(c.Writer); err != nil {
		c.String(http.StatusInternalServerError, "Gagal menyimpan file Excel")
	}
}
