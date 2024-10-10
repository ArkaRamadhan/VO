import axios from "axios";

const API_URL = "http://localhost:8080/booking-rapat";

export function getRequest(callback) {
  return axios
    .get(`${API_URL}?status=pending`)
    .then((response) => {
      callback(response.data.booking);
    })
    .catch((error) => {
      throw new Error(`Gagal mengambil data. Alasan: ${error.message}`);
    });
}

export function updateRequest(id, data) {
  return axios
    .put(`${API_URL}/${id}`, data)
    .then((response) => {
      return response.data; // Pastikan mengembalikan seluruh response untuk mendapatkan konflik
    })
    .catch((error) => {
      throw new Error(`Gagal mengubah data. Alasan: ${error.message}`);
    });
}

export function deleteRequest(id) {
  return axios
    .delete(`${API_URL}/${id}`)
    .then((response) => {
      return response.data.booking;
    })
    .catch((error) => {
      throw new Error(`Gagal menghapus data. Alasan: ${error.message}`);
    });
}
