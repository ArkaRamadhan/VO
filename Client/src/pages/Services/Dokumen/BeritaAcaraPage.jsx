import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addBeritaAcara,
  deleteBeritaAcara,
  getBeritaAcaras,
  updateBeritaAcara,
  getBeritaAcaraShow,
} from "../../../../API/Dokumen/BeritaAcara.service";
import { useToken } from "../../../context/TokenContext";
import { Modal } from "flowbite-react";
import { FormatDate } from "../../../Utilities/FormatDate";

export function BeritaAcaraPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "tanggal", label: "Tanggal", type: "date", required: false },
      {
        name: "no_surat",
        label: "Nomor Surat",
        type: "select",
        options: ["ITS-SAG", "ITS-ISO"], // Hanya kategori
        required: false,
      },
      { name: "perihal", label: "Perihal", type: "text", required: false },
      { name: "pic", label: "Pic", type: "text", required: false },
    ],
    services: "Berita Acara",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [selectedBeritaAcara, setSelectedBeritaAcara] = useState(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const handleShow = async (id) => {
    try {
      console.log("Fetching berita acara with ID:", id); // Log ID yang akan diambil
      const beritaAcara = await getBeritaAcaraShow(id);
      console.log("Fetched berita acara:", beritaAcara); // Log data berita acara yang diambil
      setSelectedBeritaAcara(beritaAcara);
      setIsModalOpen(true);
    } catch (error) {
      console.error("Error fetching berita acara:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getBeritaAcaras}
          set={addBeritaAcara}
          update={updateBeritaAcara}
          remove={deleteBeritaAcara}
          excel={{
            exportThis: "exportBeritaAcara",
            updateThis: "updateBeritaAcara",
            import: "uploadBeritaAcara",
          }}
          InfoColumn={true}
          UploadArsip={{
            get: "filesBeritaAcara",
            upload: "uploadFileBeritaAcara",
            download: "downloadBeritaAcara",
            delete: "deleteBeritaAcara",
          }}
          OnShow={handleShow} // Tambahkan prop onShow
        />
        {/* End Table */}
      </div>
      {/* Show Modal */}
      <Modal show={isModalOpen} onClose={() => setIsModalOpen(false)}>
        <Modal.Header>
          Berita Acara
        </Modal.Header>
        <Modal.Body>
          {selectedBeritaAcara && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedBeritaAcara.tanggal)}</p>
              <p className="font-semibold">Nomor Surat:</p>
              <p>{selectedBeritaAcara.no_surat}</p>
              <p className="font-semibold">Perihal:</p>
              <p>{selectedBeritaAcara.perihal}</p>
              <p className="font-semibold">Pic:</p>
              <p>{selectedBeritaAcara.pic}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
      {/* End Show Modal */}
    </App>
  );
}
