import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getArsip,
  addArsip,
  deleteArsip,
  updateArsip,
  getArsipShow,
} from "../../../../API/DataInformasi/Arsip.service";
import { useToken } from "../../../context/TokenContext";
import { Modal } from "flowbite-react";
import { FormatDate } from "../../../Utilities/FormatDate";

export function ArsipPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { name: "no_arsip", label: "No Arsip", type: "text", required: false },
      {
        name: "jenis_dokumen",
        label: "Jenis Dokumen",
        type: "text",
        required: false,
      },
      { name: "no_dokumen", label: "No Dokumen", type: "text", required: false }, // Diubah dari "From" menjadi "No Dokumen"
      { name: "perihal", label: "Perihal", type: "text", required: false }, // Tambahkan field ini
      { name: "no_box", label: "No Box", type: "text", required: false }, // Tambahkan field ini
      {
        name: "keterangan",
        label: "Keterangan",
        type: "text",
        required: false,
      }, // Tambahkan field ini
      {
        name: "tanggal_dokumen",
        label: "Tanggal Dokumen",
        type: "date",
        required: false,
      }, // Tambahkan field ini
      {
        name: "tanggal_penyerahan",
        label: "Tanggal Penyerahan",
        type: "date",
        required: false,
      }, // Tambahkan field ini
    ],
    services: "Arsip",
  });

  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedArsip, setSelectedArsip] = useState({});


  const handleShow = async (id) => {
    try {
      const Arsip = await getArsipShow(id);
      setSelectedArsip(Arsip);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching arsip:", error);
    }
  };

  const {  no_arsip, jenis_dokumen, no_dokumen, keterangan, tanggal_dokumen, tanggal_penyerahan} = selectedArsip 

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getArsip}
          set={addArsip}
          update={updateArsip}
          remove={deleteArsip}
          excel={{
            exportThis: "exportArsip",
            import: "uploadArsip",
          }}
          InfoColumn={true}
          UploadArsip={{
            get: "files",
            upload: "upload",
            download: "download",
            delete: "delete",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>

      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Arsip</Modal.Header>
        <Modal.Body>
          {selectedArsip && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">No Arsip:</p>
              <p>{no_arsip}</p>
              <p className="font-semibold">Jenis Dokumen:</p>
              <p>{jenis_dokumen}</p>
              <p className="font-semibold">No Dokumen:</p>
              <p>{no_dokumen}</p>
              <p className="font-semibold">Keterangan:</p>
              <p>{keterangan}</p>
              <p className="font-semibold">Tanggal Dokumen:</p>
              <p>{FormatDate(tanggal_dokumen)}</p>
              <p className="font-semibold">Tanggal Penyerahan:</p>
              <p>{FormatDate(tanggal_penyerahan)}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
    </App>
  );
}
