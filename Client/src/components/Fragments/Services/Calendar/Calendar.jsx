import { useState, useEffect } from "react";
import Swal from "sweetalert2";
import { Button, Label, Modal, TextInput } from "flowbite-react";
import { ColorPick } from "../../../../Utilities/ColorPick";
import FullCalendar from "@fullcalendar/react";
import { format } from "date-fns";
import idLocale from "date-fns/locale/id";
// import { formatDate } from "@fullcalendar/core";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import interactionPlugin from "@fullcalendar/interaction";
import listPlugin from "@fullcalendar/list";
import { Excel } from "../../../../Utilities/Excel"; // Import komponen Excel
import { useToken } from "../../../../context/TokenContext";
import { jwtDecode } from "jwt-decode";

export const Calendar = ({ view, get, add, remove, excel }) => {
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [formData, setFormData] = useState({});
  const [currentEvents, setCurrentEvents] = useState([]);

  const { token } = useToken(); // Ambil token dari context
  let userRole = "";
  if (token) {
    const decoded = jwtDecode(token);
    userRole = decoded.role;
  }

  const onCloseFormModal = () => {
    setFormModalOpen(false);
    setFormData({});
  };

  // Fetch events
  useEffect(() => {
    get((data) => {
      setCurrentEvents(data);
    });
  }, []);

  const handleFormChange = (e) => {
    const { name, value } = e.target;
    setFormData((prevData) => ({ ...prevData, [name]: value }));
  };

  const handleFormSubmit = async (e) => {
    e.preventDefault();
    const newEvent = {
      title: formData.title,
      start: formData.start,
      end: formData.end,
      allDay: formData.allDay,
      color: formData.color,
    };
    try {
      const response = await add(newEvent);
      setCurrentEvents([response, ...currentEvents]);
      onCloseFormModal();

      // Cek status event dan tampilkan pesan jika pending
      if (response.status === "pending") {
        Swal.fire({
          icon: "info",
          title: "Perhatian!",
          text: "Jadwal ini akan dipending karena bentrok dengan jadwal lain.",
          showConfirmButton: true,
        }).then((result) => {
          if (result.isConfirmed) {
            window.location.reload(); // Refresh halaman setelah user mengklik OK
          }
        });
      }
    } catch (error) {
      Swal.fire({
        icon: "error",
        title: "Gagal!",
        text: "Error saat menyimpan data: " + error.message,
        showConfirmButton: false,
        timer: 1500,
      });
    }
  };

  // Handle date click to add new event
  const handleDateClick = (selected) => {
    setFormModalOpen(true);
    setFormData({
      title: "",
      start: selected.startStr,
      end: selected.endStr,
      allDay: selected.allDay,
      color: "#4285f4",
    });
  };

  // Handle event click to delete event
  const handleEventClick = (selected) => {
    Swal.fire({
      title: "Apakah Anda yakin?",
      text: `Anda akan menghapus data ${selected.event.title}?`,
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Ya, saya yakin",
      cancelButtonText: "Batal",
    }).then((result) => {
      if (result.isConfirmed) {
        remove(selected.event.id)
          .then(() => {
            get((data) => {
              setCurrentEvents(data); // Pastikan data adalah array
            });
          })
          .catch((error) => {
            Swal.fire({
              icon: "error",
              title: "Gagal!",
              text: "Error saat hapus data: " + error.message,
              showConfirmButton: false,
              timer: 1500,
            });
          });
      }
    });
  };

  return (
    <div className="grid grid-cols-2fr">
      <div className="bg-gray-50 p-[15px] rounded w-[200px] max-h-[85vh] overflow-auto">
        <h2 className="text-xl mt-0 mb-2 font-bold flex justify-between">
          {currentEvents.length} Jadwal
          {userRole === "admin" && excel && (
            <Excel
              linkExportThis={excel.exportThis}
              linkUpdateThis={excel.updateThis}
              importExcel={excel.import}
            />)
          }
        </h2>
        <div className="flex flex-col gap-2">
          {currentEvents.length === 0 ? (
            <div className="ring-2 ring-blue-200 rounded p-2">
              <p>
                Belum ada jadwal yang tersedia, tambahkan jadwal baru di menu
                Calendar.
              </p>
            </div>
          ) : (
            currentEvents.map((event) => {
              return (
                <div
                  key={event.id}
                  className="ring-2 ring-slate-200 overflow-auto scrollbar-thin scrollbar-thumb-gray-400 scrollbar-track-gray-200 grow shadow py-2 px-3 rounded"
                  style={{ backgroundColor: event.color }}
                >
                  <div className="font-bold text-white">{event.title}</div>
                  <div className="text-white">
                    {format(event.start, "dd MMMM", {
                      locale: idLocale,
                    })}
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
      <div className="mx-3 max-h-[85vh] overflow-auto">
        <FullCalendar
          locale={idLocale}
          titleFormat={{
            year: "numeric",
            day: "numeric",
            month: "long",
          }}
          slotLabelFormat={{
            hour: "2-digit",
            minute: "2-digit",
            hour12: false,
          }}
          plugins={[
            dayGridPlugin,
            timeGridPlugin,
            interactionPlugin,
            listPlugin,
          ]}
          headerToolbar={{
            left: "prev,next today",
            center: "title",
            right: "dayGridMonth,timeGridWeek,timeGridDay,listMonth,exportExcel",
          }}
          buttonText={{
            today: "Hari Ini",
            month: "Bulan",
            week: "Minggu",
            day: "Hari",
            list: "Agenda",
          }}
          initialView={view}
          editable={true}
          selectable={true}
          selectMirror={true}
          dayMaxEvents={true}
          events={currentEvents}
          select={handleDateClick}
          eventClick={userRole === "admin" ? handleEventClick : undefined}
        />
      </div>


      {/* ModalForm */}
      <Modal show={formModalOpen} size="xl" onClose={onCloseFormModal} popup>
        <Modal.Header />
        <Modal.Body>
          <form onSubmit={handleFormSubmit}>
            <div className="flex flex-col gap-4">
              <div className="flex flex-col">
                <Label htmlFor="title" value="Title" />
                <TextInput
                  id="title"
                  name="title"
                  type="text"
                  className="block py-2.5 px-0 w-full text-sm text-gray-900 bg-transparent border-0 appearance-none  focus:outline-none focus:ring-0 focus:border-blue-600 peer"
                  placeholder="masukan event"
                  value={formData.title}
                  onChange={handleFormChange}
                  required
                />
              </div>
              <div className="flex flex-col gap-2 justify-start">
                <Label htmlFor="color" value="Color" />
                <ColorPick
                  colors={[
                    {
                      id: "blue",
                      label: "Blue",
                      value: "#4285f4",
                      checked: true,
                    },
                    { id: "red", label: "Red", value: "#db4437" },
                    { id: "yellow", label: "Yellow", value: "#fbbc05" },
                    { id: "green", label: "Green", value: "#0f9d58" },
                    { id: "teal", label: "Teal", value: "#00bfa5" },
                    { id: "purple", label: "Purple", value: "#9c27b0" },
                    { id: "pink", label: "Pink", value: "#e91e63" },
                  ]}
                  name="color"
                  value={formData.color}
                  onChange={handleFormChange}
                  className="mb-2 p-[2px]"
                />
              </div>
              <Button className="col-span-4" type="submit">
                Simpan
              </Button>
            </div>
          </form>
        </Modal.Body>
      </Modal>
      {/* endModalForm */}
    </div>
  );
};
