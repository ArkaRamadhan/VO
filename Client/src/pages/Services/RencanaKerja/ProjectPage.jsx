import React, { useState } from "react";
import App from "../../../components/Layouts/App";
import { ReusableTable } from "../../../components/Fragments/Services/ReusableTable";
import { jwtDecode } from "jwt-decode";
import {
  getProjects,
  addProject,
  updateProject,
  deleteProject,
  getProjectShow,
} from "../../../../API/RencanaKerja/Project.service";
import { useToken } from "../../../context/TokenContext";
import { Modal, Button } from "flowbite-react"; // Import Modal dan Button dari flowbite-react
import { FormatDate } from "../../../Utilities/FormatDate";

export function ProjectPage() {
  const [formConfig, setFormConfig] = useState({
    fields: [
      {
        name: "group",
        label: "Group",
        type: "select",
        options: ["ITS-ISO", "ITS-SAG"],
        required: true,
        visible: false,
      },
      {
        name: "infra_type",
        label: "Infrastructure Type",
        type: "select",
        options: ["SOF", "HAR", "SER"],
        required: true,
        visible: false,
      },
      {
        name: "budget_type",
        label: "Type Anggaran",
        type: "select",
        options: ["RBB", "NRBB"],
        required: true,
        visible: false,
      },
      {
        name: "type",
        label: "Type",
        type: "select",
        options: ["A", "B"],
        required: true,
        visible: false,
      },
      {
        name: "jenis_pengadaan",
        label: "Jenis Pengadaan",
        type: "text",
        required: false,
      },
      {
        name: "nama_pengadaan",
        label: "Nama Pengadaan",
        type: "text",
        required: false,
      },
      {
        name: "div_inisiasi",
        label: "Div Inisisasi",
        type: "text",
        required: false,
      },
      { name: "bulan", label: "Bulan", type: "date", required: false },
      {
        name: "sumber_pendanaan",
        label: "Sumber Pendanaan",
        type: "text",
        required: false,
      },
      { name: "anggaran", label: "Anggaran", type: "number", required: false },
      { name: "no_izin", label: "No Izin Prinsip", type: "text", required: false },
      {
        name: "tanggal_izin",
        label: "Tanggal Izin",
        type: "date",
        required: false,
      },
      {
        name: "tanggal_tor",
        label: "Tanggal Tor",
        type: "date",
        required: false,
      },
      { name: "pic", label: "PIC", type: "text", required: false },
    ],
    services: "Project",
  });
  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const [isShowModalOpen, setIsShowModalOpen] = useState(false);
  const [selectedProject, setSelectedProject] = useState(null);

  const handleShow = async (id) => {
    try {
      const project = await getProjectShow(id)
      setSelectedProject(project);
      setIsShowModalOpen(true);
    } catch (error) {
      console.error("Error fetching Project:", error);
    }
  };

  return (
    <App services={formConfig.services}>
      <div className="overflow-auto">
        {/* Table */}
        <ReusableTable
          formConfig={formConfig}
          setFormConfig={setFormConfig}
          get={getProjects}
          set={addProject}
          update={updateProject}
          remove={deleteProject}
          excel={{
            exportThis: "exportProject",
            updateThis: "updateProject",
            import  : "uploadProject",
          }}  
          InfoColumn={true}
          KodePj={true}
          UploadArsip={{
            get: "filesProject",
            upload: "uploadFileProject",
            download: "downloadProject",
            delete: "deleteProject",
          }}
          OnShow={handleShow}
        />
        {/* End Table */}
      </div>

      {/* Modal */}
      <Modal show={isShowModalOpen} onClose={() => setIsShowModalOpen(false)}>
        <Modal.Header>Detail Project</Modal.Header>
        <Modal.Body>
          {selectedProject && (
            <div className="grid grid-cols-2 gap-4">
              <p className="font-semibold">Kode Project :</p>
              <p>{selectedProject.kode_project}</p>
              <p className="font-semibold">Jenis Pengadaan:</p>
              <p>{selectedProject.jenis_pengadaan}</p>
              <p className="font-semibold">Nama Pengadaan:</p>
              <p>{selectedProject.nama_pengadaan}</p>
              <p className="font-semibold">Div Inisisasi:</p>
              <p>{selectedProject.div_inisiasi}</p>
              <p className="font-semibold">Bulan:</p>
              <p>
                {selectedProject.bulan 
                  ? FormatDate(new Date(selectedProject.bulan)) 
                  : "N/A"}
              </p>
              <p className="font-semibold">Sumber Pendanaan:</p>
              <p>{selectedProject.sumber_pendanaan}</p>
              <p className="font-semibold">Anggaran:</p>
              <p>{selectedProject.anggaran}</p>
              <p className="font-semibold">No Izin Prinsip:</p>
              <p>{selectedProject.no_izin}</p>
              <p className="font-semibold">Tanggal Izin:</p>
              <p>
                {selectedProject.tanggal_izin 
                  ? FormatDate(new Date(selectedProject.tanggal_izin)) 
                  : "N/A"}
              </p>
              <p className="font-semibold">Tanggal Tor:</p>
              <p>
                {selectedProject.tanggal_tor 
                  ? FormatDate(new Date(selectedProject.tanggal_tor)) 
                  : "N/A"}
              </p>
              <p className="font-semibold">PIC:</p>
              <p>{selectedProject.pic}</p>
            </div>
          )}
        </Modal.Body> 
      </Modal>
      {/* End Modal */}
    </App>
  );
}
