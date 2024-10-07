import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addSk,
  deleteSk,
  getSks,
  updateSk,
  getSkShow
} from "../../../../API/Dokumen/Sk.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate";

export function SkPage() {
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
    services: "Sk",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedSk, setSelectedSk] = useState(null);

  const handleShow = async (id) => {
    try {
      const sk = await getSkShow(id);
      setSelectedSk(sk);
      console.log(sk);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching SK:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getSks}
          set={addSk}
          update={updateSk}
          remove={deleteSk}
          excel={{
            exportThis: "exportSk",
            updateThis: "updateSk",
            import  : "uploadSk",
          }}
          InfoColumn={true}
          UploadArsip={{
            get: "filesSk",
            upload: "uploadFileSk",
            download: "downloadSk",
            delete: "deleteSk",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>
      {/* Show Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail SK</Modal.Header>
        <Modal.Body>
          {selectedSk && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedSk.tanggal)}</p>
              <p className="font-semibold">Nomor Surat:</p>
              <p>{selectedSk.no_surat}</p>
              <p className="font-semibold">Perihal:</p>
              <p>{selectedSk.perihal}</p>
              <p className="font-semibold">PIC:</p>
              <p>{selectedSk.pic}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
      {/* End Show Modal */}
    </App>
  );
}
