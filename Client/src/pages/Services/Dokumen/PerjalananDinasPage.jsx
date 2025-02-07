import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getPerdins,
  addPerdin,
  deletePerdin,
  updatePerdin,
  getPerdinShow,
} from "../../../../API/Dokumen/PerjalananDinas.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate";

export function PerdinPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "no_perdin", label: "No Perdin", type: "text", required: false },
      {
        name: "tanggal",
        label: "Tanggal Pengajuan",
        type: "date",
        required: false,
      },
      { name: "hotel", label: "Deskripsi", type: "text", required: false },
      { name: "transport", label: "~", type: "text", required: false },
    ],
    services: "Perjalanan Dinas",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedPerdin, setSelectedPerdin] = useState(null);

  const handleShow = async (id) => {
    try {
      const perdin = await getPerdinShow(id);
      setSelectedPerdin(perdin);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching Perdin:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Tabel */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getPerdins}
          set={addPerdin}
          update={updatePerdin}
          remove={deletePerdin}
          excel={{
            exportThis: "exportPerdin",
            updateThis: "updatePerdin",
            import  : "uploadPerdin",
          }}  
          InfoColumn={true}
          UploadArsip={{
            get: "filesPerdin",
            upload: "uploadFilePerdin",
            download: "downloadPerdin",
            delete: "deletePerdin",
          }}
          OnShow={handleShow}
        />
        {/* End Tabel */}
      </div>
      {/* Show Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Perdin</Modal.Header>
        <Modal.Body>
          {selectedPerdin && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedPerdin.tanggal)}</p>
              <p className="font-semibold">No Perdin:</p>
              <p>{selectedPerdin.no_perdin}</p>
              <p className="font-semibold">Deskripsi:</p>
              <p>{selectedPerdin.hotel}</p>
              <p className="font-semibold">~:</p>
              <p>{selectedPerdin.transport}</p> 
            </div>
          )}
        </Modal.Body>
      </Modal>
      {/* End Show Modal */}
    </App>
  );
}
