import axios from "axios";

const API_URL = "http://localhost:8080/surat";

export function getSurats(callback) {
  return axios
    .get(`${API_URL}`)
    .then((response) => {
      callback(response.data.surat);
    })
    .catch((error) => {
      throw new Error(`Gagal mengambil data. Alasan: ${error.message}`);
    });
}

export function addSurat(data) {
  const { username, ...rest } = data;
  return axios
    .post(`${API_URL}`, { ...rest }) // Tambahkan info
    .then((response) => {
      return response.data.surat;
    })
    .catch((error) => {
      throw new Error(`Gagal menambahkan data. Alasan: ${error.message}`);
    });
}

export function updateSurat(id, data) {
  const { username, ...rest } = data;
  return axios
    .put(`${API_URL}/${id}`, { ...rest }) // Kirim username sebagai bagian dari request
    .then((response) => {
      return response.data.surat;
    })
    .catch((error) => {
      throw new Error(`Gagal mengubah data. Alasan: ${error.message}`);
    });
}

export function getSuratShow(id) {
  return axios
    .get(`${API_URL}/${id}`)
    .then((response) => {
      return response.data.surat;
    })
    .catch((error) => {
      throw new Error(`Gagal mengambil data. Alasan: ${error.message}`);
    });
}

export function deleteSurat(id) {
  return axios
    .delete(`${API_URL}/${id}`)
    .then((response) => {
      return response.data.surat;
    })
    .catch((error) => {
      throw new Error(`Gagal menghapus data. Alasan: ${error.message}`);
    });
}