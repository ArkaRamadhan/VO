import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addSurat,
  deleteSurat,
  getSurats,
  updateSurat,
  getSuratShow,
} from "../../../../API/Dokumen/Surat.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate";

export function SuratPage() {
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
    services: "Surat",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedSurat, setSelectedSurat] = useState(null);

  const handleShow = async (id) => {
    try {
      const surat = await getSuratShow(id);
      setSelectedSurat(surat);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching surat:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getSurats}
          set={addSurat}
          update={updateSurat}
          remove={deleteSurat}
          excel={{
            exportThis: "exportSurat",
            import: "uploadSurat",
          }}
          OnShow={handleShow}
          InfoColumn={true}
          UploadArsip={{
            get: "filesSurat",
            upload: "uploadFileSurat",
            download: "downloadSurat",
            delete: "deleteSurat",
          }}
          filterSagIso
        />
        {/* End Table */}

      </div>
      {/* Show Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Surat</Modal.Header>
        <Modal.Body>
          {selectedSurat && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedSurat.tanggal)}</p>
              <p className="font-semibold">Nomor Surat:</p>
              <p>{selectedSurat.no_surat}</p>
              <p className="font-semibold">Perihal:</p>
              <p>{selectedSurat.perihal}</p>
              <p className="font-semibold">Pic:</p>
              <p>{selectedSurat.pic}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
      {/* End Show Modal */}
    </App>
  );
}
