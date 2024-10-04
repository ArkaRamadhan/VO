import React, { useState, useEffect } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getPerdins,
  addPerdin,
  deletePerdin,
  updatePerdin,
} from "../../../../API/Dokumen/PerjalananDinas.service";
import { useToken } from "../../../context/TokenContext";

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

  const { token } = useToken();
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [defaultValues, setDefaultValues] = useState({
    no_perdin: "PD-ITS",
    tanggal: "",
    hotel: "",
    transport: "",
  });

  useEffect(() => {
    // Set default values in the formConfig
    setFormConfig((prevConfig) => ({
      ...prevConfig,
      defaultValues,
    }));
  }, []); // Hanya panggil saat komponen di-mount

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getPerdins}
          set={addPerdin}
          update={updatePerdin}
          remove={deletePerdin}
          excel
          ExportExcel="exportPerdin"
          UpdateExcel="updatePerdin"
          ImportExcel="uploadPerdin"
          InfoColumn={true}
          UploadArsip={{
            get: "filesPerdin",
            upload: "uploadFilePerdin",
            download: "downloadPerdin",
            delete: "deletePerdin",
          }}
        />
      </div>
    </App>
  );
}
