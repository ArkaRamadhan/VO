// Format tanggal ke string dengan format indonesia atau bulan/tahun
export const FormatDate = (dateString, formatType = "indonesia") => {
  const date = new Date(dateString);

  if (isNaN(date.getTime())) {
    return ""; // Return empty string if date is invalid
  }

  if (formatType === "bulanTahun") {
    const month = date.getMonth() + 1; // getMonth() returns 0-11
    const year = date.getFullYear(); // Use full year
    return `${month.toString().padStart(2, '0')}/${year}`; // Format <bulan>/<tahun>
  } else {
    const months = [
      "Januari",
      "Februari",
      "Maret",
      "April",
      "Mei",
      "Juni",
      "Juli",
      "Agustus",
      "September",
      "Oktober",
      "November",
      "Desember",
    ];
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
  }
};
