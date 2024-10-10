import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addMeeting,
  deleteMeeting,
  getMeetings,
  updateMeeting,
  getMeetingShow,
} from "../../../../API/KegiatanProses/Meeting.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate";

export function MeetingPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      { 
        name: "task",
        label: "Task",
        type: "text",
        required: true,
      },
      { name: "tindak_lanjut", label: "Tindak Lanjut", type: "text", required: true },
      {
        name: "status",
        label: "Status",
        type: "select",
        options: ["Done", "On Progress", "Cancel"],
        required: true,
      },
      { name: "update_pengerjaan", label: "Update Pengerjaan", type: "text", required: false },
      { name: "pic", label: "Pic", type: "text", required: true },
      { name: "tanggal_target", label: "Tanggal Target", type: "date", required: true },
      { name: "tanggal_actual", label: "Tanggal Actual", type: "date", required: true },
    ],
    services: "Meeting",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedMeeting, setSelectedMeeting] = useState(null);

  const handleShow = async (id) => {
    try {
      const meeting = await getMeetingShow(id)
      setSelectedMeeting(meeting);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching Meeting:", error);
    }
  };
  
  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getMeetings}
          set={addMeeting}
          update={updateMeeting}
          remove={deleteMeeting}
          excel={{
            exportThis: "exportMeeting",
            import: "uploadMeeting",
          }}
          InfoColumn={true}
          StatusColumn={true}
          OnShow={handleShow}
          UploadArsip={{
            get: "filesMeeting",
            upload: "uploadFileMeeting",
            download: "downloadMeeting",
            delete: "deleteMeeting",
          }}
        />
        {/* End Table */}
      </div>

      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>
          Meeting Detail
        </Modal.Header>
        <Modal.Body>
          {selectedMeeting && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Task :</p>
              <p>{selectedMeeting.task}</p>
              <p className="font-semibold">Tindak Lanjut :</p>
              <p>{selectedMeeting.tindak_lanjut}</p>
              <p className="font-semibold">Status :</p>
              <p>{selectedMeeting.status}</p>
              <p className="font-semibold">Update Pengerjaan :</p>
              <p>{selectedMeeting.update_pengerjaan}</p>
              <p className="font-semibold">Pic :</p>
              <p>{selectedMeeting.pic}</p>
              <p className="font-semibold">Tanggal Target :</p>
              <p>{FormatDate(selectedMeeting.tanggal_target)}</p>
              <p className="font-semibold">Tanggal Actual :</p>
              <p>{FormatDate(selectedMeeting.tanggal_actual)}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
    </App>
  );
}
