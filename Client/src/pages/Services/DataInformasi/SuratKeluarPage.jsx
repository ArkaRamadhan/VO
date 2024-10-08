import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getSuratKeluars,
  addSuratKeluar,
  deleteSuratKeluar,
  updateSuratKeluar,
  getSuratKeluarShow,
} from "../../../../API/DataInformasi/SuratKeluar.service";
import { useToken } from "../../../context/TokenContext";
import { Modal } from "flowbite-react";
import { FormatDate } from "../../../Utilities/FormatDate";

export function SuratKeluarPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "no_surat", label: "No Surat", type: "text", required: false },
      { name: "title", label: "Title Of Letter", type: "text", required: false },
      { name: "from", label: "From", type: "text", required: false },
      { name: "pic", label: "PIC", type: "text", required: false },
      { name: "tanggal", label: "Date Issue", type: "date", required: false },
    ],
    services: "Surat Keluar",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedSuratKeluar, setSelectedSuratKeluar] = useState(null);

  const handleShow = async (id) => {
    try {
      const SuratKeluar = await getSuratKeluarShow(id);
      setSelectedSuratKeluar(SuratKeluar);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching surat keluar:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getSuratKeluars}
          set={addSuratKeluar}
          update={updateSuratKeluar}
          remove={deleteSuratKeluar}
          excel
          ExportExcel="exportSuratKeluar"
          UpdateExcel="updateSuratKeluar"
          ImportExcel="uploadSuratKeluar"
          InfoColumn={true}
          UploadArsip={{
            get: "filesSuratKeluar",
            upload: "uploadFileSuratKeluar",
            download: "downloadSuratKeluar",
            delete: "deleteSuratKeluar",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>
      {/* Show Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Surat Keluar</Modal.Header>
        <Modal.Body>
          {selectedSuratKeluar && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">No Surat:</p>
              <p>{selectedSuratKeluar.no_surat}</p>
              <p className="font-semibold">Title:</p>
              <p>{selectedSuratKeluar.title}</p>
              <p className="font-semibold">From:</p>
              <p>{selectedSuratKeluar.from}</p>
              <p className="font-semibold">PIC:</p>
              <p>{selectedSuratKeluar.pic}</p>
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedSuratKeluar.tanggal)}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
    </App>
  );
}
