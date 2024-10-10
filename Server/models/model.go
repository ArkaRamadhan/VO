package models

import (
	"encoding/json"
	"time"

	"gorm.io/gorm"
)

// model for user
type User struct {
	gorm.Model
	Username string
	Email    string
	Password string
	Role     string
	Info     string
}

type UserToken struct {
	ID     uint   `gorm:"primaryKey"`
	UserID uint   `gorm:"not null"`
	Token  string `gorm:"not null"`
	Expiry time.Time
}

// model for memo
type Memo struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	Tanggal   *time.Time `json:"tanggal"`
	NoMemo    *string    `json:"no_memo"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

// MarshalJSON menyesuaikan serialisasi JSON untuk struct Memo
func (i *Memo) MarshalJSON() ([]byte, error) {
	type Alias Memo
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type BeritaAcara struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *BeritaAcara) MarshalJSON() ([]byte, error) {
	type Alias BeritaAcara
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type Surat struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *Surat) MarshalJSON() ([]byte, error) {
	type Alias Surat
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type Sk struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *Sk) MarshalJSON() ([]byte, error) {
	type Alias Sk
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

// list meeting
type MeetingSchedule struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	Hari      *string    `json:"hari"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Waktu     *string    `json:"waktu"`
	Selesai   *string    `json:"selesai"`
	Tempat    *string    `json:"tempat"`
	Pic       *string    `json:"pic"`
	Status    *string    `json:"status"`
	CreateBy  string     `json:"create_by"`
	Color     string     `json:"color"`
}

func (i *MeetingSchedule) MarshalJSON() ([]byte, error) {
	type Alias MeetingSchedule
	tanggalFormatted := i.Tanggal.Format("2006-01-02")
	return json.Marshal(&struct {
		Tanggal *string `json:"tanggal"`
		*Alias
	}{
		Tanggal: &tanggalFormatted,
		Alias:   (*Alias)(i),
	})
}

// model for meeting
type Meeting struct {
	ID               uint       `gorm:"primaryKey"`
	CreatedAt        *time.Time `gorm:"autoCreateTime"`
	UpdatedAt        *time.Time `gorm:"autoUpdateTime"`
	Task             *string    `json:"task"`
	TindakLanjut     *string    `json:"tindak_lanjut"`
	Status           *string    `json:"status"`
	UpdatePengerjaan *string    `json:"update_pengerjaan"`
	Pic              *string    `json:"pic"`
	TanggalTarget    *time.Time `json:"tanggal_target"`
	TanggalActual    *time.Time `json:"tanggal_actual"`
	CreateBy         string     `json:"create_by"`
}

func (i *Meeting) MarshalJSON() ([]byte, error) {
	type Alias Meeting
	tanggalTargetFormatted := i.TanggalTarget.Format("2006-01-02")
	tanggalActualFormatted := i.TanggalActual.Format("2006-01-02")
	return json.Marshal(&struct {
		TanggalTarget *string `json:"tanggal_target"`
		TanggalActual *string `json:"tanggal_actual"`
		*Alias
	}{
		TanggalTarget: &tanggalTargetFormatted,
		TanggalActual: &tanggalActualFormatted,
		Alias:         (*Alias)(i),
	})
}

// model for perdin
type Perdin struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoPerdin  *string    `json:"no_perdin"`
	Tanggal   *time.Time `json:"tanggal"`
	Hotel     *string    `json:"hotel"`
	Transport *string    `json:"transport"`
	CreateBy  string     `json:"create_by"`
}

func (i *Perdin) MarshalJSON() ([]byte, error) {
	type Alias Perdin
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}


// model for project
type Project struct {
	ID              uint       `gorm:"primaryKey"`
	CreatedAt       *time.Time `gorm:"autoCreateTime"`
	UpdatedAt       *time.Time `gorm:"autoUpdateTime"`
	KodeProject     *string    `json:"kode_project"`
	JenisPengadaan  *string    `json:"jenis_pengadaan"`
	NamaPengadaan   *string    `json:"nama_pengadaan"`
	DivInisiasi     *string    `json:"div_inisiasi"`
	Bulan           *time.Time `json:"bulan"`
	SumberPendanaan *string    `json:"sumber_pendanaan"`
	Anggaran        *string    `json:"anggaran"`
	NoIzin          *string    `json:"no_izin"`
	TanggalIzin     *time.Time `json:"tanggal_izin"`
	TanggalTor      *time.Time `json:"tanggal_tor"`
	Pic             *string    `json:"pic"`
	Group           *string    `json:"group"`
	InfraType       *string    `json:"infra_type"`
	BudgetType      *string    `json:"budget_type"`
	Type            *string    `json:"type"`
	CreateBy        string     `json:"create_by"`
}

func (p *Project) MarshalJSON() ([]byte, error) {
	type Alias Project
	var tanggalIzinFormatted, tanggalTorFormatted, bulanFormatted string

	// Cek TanggalIzin
	if p.TanggalIzin == nil {
		tanggalIzinFormatted = ""
	} else {
		tanggalIzinFormatted = p.TanggalIzin.Format("2006-01-02")
	}

	// Cek TanggalTor
	if p.TanggalTor == nil {
		tanggalTorFormatted = ""
	} else {
		tanggalTorFormatted = p.TanggalTor.Format("2006-01-02")
	}

	// Cek Bulan
	if p.Bulan == nil {
		bulanFormatted = ""
	} else {
		bulanFormatted = p.Bulan.Format("01/06")
	}

	return json.Marshal(&struct {
		TanggalIzin string `json:"tanggal_izin"`
		TanggalTor  string `json:"tanggal_tor"`
		Bulan       string `json:"bulan"`
		*Alias
	}{
		TanggalIzin: tanggalIzinFormatted,
		TanggalTor:  tanggalTorFormatted,
		Bulan:       bulanFormatted,
		Alias:       (*Alias)(p),
	})
}

// model jadwal-rapat
type Notification struct {
	ID       uint      `gorm:"primaryKey" json:"id"`
	Title    string    `json:"title"`
	Start    time.Time `json:"start"`
	Category string    `json:"category"`
}

type BookingRapat struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"` // Tambahkan field ini untuk warna
	Status string `json:"status"`
}

func (BookingRapat) TableName() string {
	return "booking_rapats"
}

type JadwalRapat struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"`
}

func (JadwalRapat) TableName() string {
	return "jadwal_rapats"
}

type JadwalCuti struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"` // Tambahkan field ini untuk warna
}

type TimelineProject struct {
	ID         uint   `gorm:"primaryKey" json:"id"`
	Start      string `json:"start"`
	End        string `json:"end"`
	ResourceId int    `json:"resourceId"` // Ubah tipe data dari string ke int
	Title      string `json:"title"`
	BgColor    string `json:"bgColor"`
}

func (TimelineProject) TableName() string {
	return "timeline_projects"
}

type ResourceProject struct {
	ID       uint   `gorm:"primaryKey" json:"id"`
	Name     string `json:"name"`
	ParentID uint   `json:"parent_id"`
}

type TimelineDesktop struct {
	ID         uint   `gorm:"primaryKey" json:"id"`
	Start      string `json:"start"`
	End        string `json:"end"`
	ResourceId int    `json:"resourceId"` // Ubah tipe data dari string ke int
	Title      string `json:"title"`
	BgColor    string `json:"bgColor"`
}

func (TimelineDesktop) TableName() string {
	return "timeline_desktops"
}

type ResourceDesktop struct {
	ID       uint   `gorm:"primaryKey" json:"id"`
	Name     string `json:"name"`
	ParentID uint   `json:"parent_id"`
}

// model for suratMasuk
type SuratMasuk struct {
	ID         uint       `gorm:"primaryKey"`
	CreatedAt  *time.Time `gorm:"autoCreateTime"`
	UpdatedAt  *time.Time `gorm:"autoUpdateTime"`
	NoSurat    *string    `json:"no_surat"`
	Title      *string    `json:"title"`
	RelatedDiv *string    `json:"related_div"`
	DestinyDiv *string    `json:"destiny_div"`
	Tanggal    *time.Time `json:"tanggal"`
	CreateBy   string     `json:"create_by"`
}

func (i *SuratMasuk) MarshalJSON() ([]byte, error) {
	type Alias SuratMasuk
	if i.Tanggal == nil {
		// Handle jika Tanggal nil
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau format default yang diinginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

// model for suratKeluar
type SuratKeluar struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Title     *string    `json:"title"`
	From      *string    `json:"from"`
	Pic       *string    `json:"pic"`
	Tanggal   *time.Time `json:"tanggal"`
	CreateBy  string     `json:"create_by"`
}

func (i *SuratKeluar) MarshalJSON() ([]byte, error) {
	type Alias SuratKeluar
	tanggalFormatted := i.Tanggal.Format("2006-01-02")
	return json.Marshal(&struct {
		Tanggal *string `json:"tanggal"`
		*Alias
	}{
		Tanggal: &tanggalFormatted,
		Alias:   (*Alias)(i),
	})
}

type Arsip struct {
	gorm.Model
	NoArsip           *string    `json:"no_arsip"`
	JenisDokumen      *string    `json:"jenis_dokumen"`
	NoDokumen         *string    `json:"no_dokumen"`
	TanggalDokumen    *time.Time `json:"tanggal_dokumen"`
	Perihal           *string    `json:"perihal"`
	NoBox             *string    `json:"no_box"`
	TanggalPenyerahan *time.Time `json:"tanggal_penyerahan"`
	Keterangan        *string    `json:"keterangan"`
	CreateBy          string     `json:"create_by"`
}

func (a *Arsip) MarshalJSON() ([]byte, error) {
	type Alias Arsip
	var tanggalDokumenFormatted, tanggalPenyerahanFormatted string

	// Cek TanggalDokumen
	if a.TanggalDokumen == nil {
		tanggalDokumenFormatted = ""
	} else {
		tanggalDokumenFormatted = a.TanggalDokumen.Format("2006-01-02")
	}

	// Cek TanggalPenyerahan
	if a.TanggalPenyerahan == nil {
		tanggalPenyerahanFormatted = ""
	} else {
		tanggalPenyerahanFormatted = a.TanggalPenyerahan.Format("2006-01-02")
	}

	return json.Marshal(&struct {
		TanggalDokumen    string `json:"tanggal_dokumen"`
		TanggalPenyerahan string `json:"tanggal_penyerahan"`
		*Alias
	}{
		TanggalDokumen:    tanggalDokumenFormatted,
		TanggalPenyerahan: tanggalPenyerahanFormatted,
		Alias:             (*Alias)(a),
	})
}

type File struct {
	ID          uint      `gorm:"primaryKey"`         // ID unik untuk file
	CreatedAt   time.Time `gorm:"autoCreateTime"`     // Timestamp saat file diunggah
	UpdatedAt   time.Time `gorm:"autoUpdateTime"`     // Timestamp untuk setiap update
	UserID      uint      `gorm:"index"`              // ID pengguna yang mengunggah file
	FilePath    string    `gorm:"not null"`           // Path lengkap di mana file disimpan
	FileName    string    `gorm:"not null"`           // Nama file asli
	ContentType string    `gorm:"not null"`           // Jenis konten file, misal 'application/pdf'
	Size        int64     `gorm:"not null"`           // Ukuran file dalam byte
}

type ConflictRequest struct {
	gorm.Model
	NewEventID uint
	OldEventID uint
	Status     string
	OldTitle   string
	NewTitle   string
	StartTime  string
	EndTime    string
	Date       time.Time
}

// TableName overrides the table name used by File to `files`, if you want to specify it explicitly
func (File) TableName() string {
	return "files"
}
