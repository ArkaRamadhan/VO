import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  addMeetingList,
  deleteMeetingList,
  getMeetingList,
  updateMeetingList,
  getMeetingListShow,
} from "../../../../API/KegiatanProses/MeetingSchedule.service";
import { useToken } from "../../../context/TokenContext";
import { Modal } from "flowbite-react";

// Komponen untuk format hari dan tanggal
const FormatHariTanggal = (tanggal) => {
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  return new Date(tanggal).toLocaleDateString('id-ID', options);
};

export function MeetingListPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      // Field baru untuk gabung hari dan tanggal
      {
        name: "tanggal",
        label: "Hari dan Tanggal",
        type: "date",
        required: true,
        render: (value) => FormatHariTanggal(value), // Tambahkan fungsi render khusus
      },
      { name: "perihal", label: "Perihal", type: "text", required: true },
      {
        name: "waktu",
        label: "Waktu",
        type: "time",
        required: false,
      },
      {
        name: "selesai",
        label: "Selesai",
        type: "time",
        required: false,
      },
      { name: "tempat", label: "Tempat", type: "text", required: false },
      {
        name: "status",
        label: "Status",
        type: "select",
        options: ["Done", "Reschedule", "On Progress", "Cancel"],
        required: true,
      },
      { name: "pic", label: "Pic", type: "text", required: true },
    ],
    services: "Meeting Schedule",
  });
  
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedMeetingList, setSelectedMeetingList] = useState(null);
  console.log(selectedMeetingList);

  const handleShow = async (id) => {
    try {
      const meetingList = await getMeetingListShow(id);
      console.log("Fetched meeting list:", meetingList);
      setSelectedMeetingList(meetingList);
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
          get={getMeetingList}
          set={addMeetingList}
          update={updateMeetingList}
          remove={deleteMeetingList}
          excel={{
            exportThis: "exportMeetingList",
            import: "uploadMeetingList",
          }}
          InfoColumn={true}
          StatusColumn={true}
          UploadArsip={{
            get: "filesMeetingList",
            upload: "uploadFileMeetingList",
            download: "downloadMeetingList",
            delete: "deleteMeetingList",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>

      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>
          <h3 className="text-xl font-semibold text-gray-900 dark:text-white">
            Meeting Detail
          </h3>
        </Modal.Header>
        <Modal.Body>
          {selectedMeetingList && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Hari dan Tanggal :</p>
              <p>
                {/* Format hari dan tanggal di sini */}
                {FormatHariTanggal(selectedMeetingList.tanggal)}
              </p>
              <p className="font-semibold">Perihal :</p>
              <p>{selectedMeetingList.perihal}</p>
              <p className="font-semibold">Waktu :</p>
              <p>{selectedMeetingList.waktu}</p>
              <p className="font-semibold">Selesai :</p>
              <p>{selectedMeetingList.selesai}</p>
              <p className="font-semibold">Tempat :</p>
              <p>{selectedMeetingList.tempat}</p>
              <p className="font-semibold">Status :</p>
              <p>{selectedMeetingList.status}</p>
              <p className="font-semibold">Pic :</p>
              <p>{selectedMeetingList.pic}</p>
            </div>
          )}
        </Modal.Body>
      </Modal>
    </App>
  );
}
