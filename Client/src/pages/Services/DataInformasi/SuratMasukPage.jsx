import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getSuratMasuks,
  addSuratMasuk,
  deleteSuratMasuk,
  updateSuratMasuk,
  getSuratMasukShow
} from "../../../../API/DataInformasi/SuratMasuk.service";
import { useToken } from "../../../context/TokenContext";
import { Modal } from "flowbite-react";
import { FormatDate } from "../../../Utilities/FormatDate";

export function SuratMasukPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "no_surat", label: "No Surat", type: "text", required: false },
      { name: "title", label: "Title Of Letter", type: "text", required: false },
      {
        name: "related_div",
        label: "Related Divisi",
        type: "text",
        required: false,
      },
      { name: "destiny_div", label: "Tujuan", type: "text", required: false },
      { name: "tanggal", label: "Date Issue", type: "date", required: false },
    ],
    services: "Surat Masuk",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedSuratMasuk, setSelectedSuratMasuk] = useState(null);

  const handleShow = async (id) => {
    try {
      const suratMasuk = await getSuratMasukShow(id);
      setSelectedSuratMasuk(suratMasuk);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching surat masuk:", error);
    }
  };
  

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getSuratMasuks}
          set={addSuratMasuk}
          update={updateSuratMasuk}
          remove={deleteSuratMasuk}
          excel={{
            exportThis : "exportSuratMasuk",
            import : "importSuratMasuk",
          }}
          InfoColumn={true}
          UploadArsip={{
            get: "filesSuratMasuk",
            upload: "uploadFileSuratMasuk",
            download: "downloadSuratMasuk",
            delete: "deleteSuratMasuk",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Surat Masuk</Modal.Header>
        <Modal.Body>
          {selectedSuratMasuk && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">No Surat:</p>
              <p>{selectedSuratMasuk.no_surat}</p>
              <p className="font-semibold">Title:</p>
              <p>{selectedSuratMasuk.title}</p>
              <p className="font-semibold">Related Divisi:</p>
              <p>{selectedSuratMasuk.related_div}</p>
              <p className="font-semibold">Destiny Divisi:</p>
              <p>{selectedSuratMasuk.destiny_div}</p>
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedSuratMasuk.tanggal)}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
    </App>
  );
}
