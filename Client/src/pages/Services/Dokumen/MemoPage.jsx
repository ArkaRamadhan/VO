import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addMemo,
  deleteMemo,
  getMemos,
  updateMemo,
  getMemoShow,
} from "../../../../API/Dokumen/MemoSag.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate"; // Import FormatDate

export function MemoPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "tanggal", label: "Tanggal", type: "date", required: true },
      {
        name: "no_memo",
        label: "Nomor Memo/Surat",
        type: "select",
        options: ["ITS-SAG", "ITS-ISO"], // Hanya kategori
        required: true,
      },
      { name: "perihal", label: "Perihal", type: "text", required: true },
      { name: "pic", label: "Pic", type: "text", required: true },
    ],
    services: "Memo",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedMemo, setSelectedMemo] = useState(null);

  const handleShow = async (id) => {
    try {
      console.log("Fetching memo with ID:", id); // Log ID yang akan diambil
      const memo = await getMemoShow(id);
      console.log("Fetched memo:", memo); // Log data memo yang diambil
      setSelectedMemo(memo);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching memo:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getMemos}
          set={addMemo}
          update={updateMemo}
          remove={deleteMemo}
          excel={{
            exportThis: "exportMemo",
            updateThis: "updateMemo",
            import: "uploadMemo",
          }}
          InfoColumn={true}
          UploadArsip={{
            get: "filesMemo",
            upload: "uploadFileMemo",
            download: "downloadMemo",
            delete: "deleteMemo",
          }}
          OnShow={handleShow} // Tambahkan prop onShow
        />
        {/* End Table */}
      </div>
      {/* Show Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Memo</Modal.Header>
        <Modal.Body>
          {selectedMemo && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Tanggal:</p>
              <p>{FormatDate(selectedMemo.tanggal)}</p>
              <p className="font-semibold">Nomor Memo/Surat:</p>
              <p>{selectedMemo.no_memo}</p>
              <p className="font-semibold">Perihal:</p>
              <p>{selectedMemo.perihal}</p>
              <p className="font-semibold">Pic:</p>
              <p>{selectedMemo.pic}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
      {/* End Show Modal */}
    </App>
  );
}