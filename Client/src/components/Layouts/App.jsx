import { Badge, Label } from "flowbite-react";
import { MdOutlineDashboard } from "react-icons/md";
import { HiOutlineClipboardDocumentList } from "react-icons/hi2";
import { GoProjectSymlink } from "react-icons/go";
import { GrPlan } from "react-icons/gr";
import { SlEnvolopeLetter } from "react-icons/sl";
import { FiUsers } from "react-icons/fi";
import { BiLogOut } from "react-icons/bi";
import { Dropdown } from "flowbite-react";
import { VscGitPullRequestGoToChanges } from "react-icons/vsc";
import { useState, useEffect, useRef } from "react";
import {
  GetNotification,
  deleteNotification,
} from "../../../API/KegiatanProses/Notification/Notification";
import { RealtimeClock, RealtimeDate } from "../../Utilities/RealTimeClock";
import { format } from "date-fns";
import idLocale from "date-fns/locale/id";
import Swal from "sweetalert2";
import { useToken } from "../../context/TokenContext";
import Sidebar, { SidebarItem, SidebarCollapse } from "../Elements/Sidebar";

const App = ({ services, children }) => {
  const { token, userDetails } = useToken(); // Ambil token dari context
  const [notification, setNotification] = useState({
    JadwalCuti: [],
    JadwalRapat: [],
    TimelineProject: [],
    TimelineWallpaperDesktop: [],
    BookingRapat: [],
  });

  const [filter, setFilter] = useState({
    JadwalCuti: false,
    JadwalRapat: true,
    BookingRapat: true,
    TimelineWallpaperDesktop: true,
    TimelineProject: true,
  });

  const [selectedNotifications, setSelectedNotifications] = useState([]);

  // State untuk menyimpan waktu event
  const [eventTime, setEventTime] = useState(null);

  const eventTimeRef = useRef(null);

  useEffect(() => {
    if (eventTime) {
      eventTimeRef.current = eventTime;
    }
  }, [eventTime]);

  useEffect(() => {
    const checkEventTime = () => {
      const now = new Date();
      const oneHourBeforeEvent = new Date(eventTimeRef.current.getTime() - 60 * 60 * 1000);

      console.log("Waktu saat ini:", now); // Log waktu saat ini
      console.log("Satu jam sebelum event:", oneHourBeforeEvent); // Log waktu satu jam sebelum event

      if (now >= oneHourBeforeEvent && now <= eventTimeRef.current) {
        console.log("Menampilkan notifikasi untuk event."); // Log ketika kondisi untuk menampilkan notifikasi terpenuhi
        Swal.fire({
          title: 'Pengingat Event',
          text: 'Event Anda akan dimulai dalam satu jam!',
          icon: 'info',
          confirmButtonText: 'Baik'
        }).then((result) => {
          if (result.isConfirmed) {
            // Hapus notifikasi setelah konfirmasi
            setNotification((prevData) => {
              const updatedNotifications = { ...prevData };
              Object.keys(updatedNotifications).forEach((category) => {
                updatedNotifications[category] = updatedNotifications[category].filter(
                  (event) => event.start !== eventTimeRef.current
                );
              });
              return updatedNotifications;
            });
          }
        });
      } else {
        console.log("Notifikasi tidak ditampilkan."); // Log ketika notifikasi tidak ditampilkan
      }
    };

    const timer = setInterval(() => {
      checkEventTime();
    }, 60000); // Periksa setiap menit

    return () => {
      console.log("Membersihkan timer."); // Log ketika membersihkan timer
      clearInterval(timer);
    };
  }, []); // useEffect ini tidak memiliki dependensi dan hanya di-set sekali

  const handleFilterChange = (category) => {
    setFilter((prevFilter) => ({
      ...prevFilter,
      [category]: !prevFilter[category],
    }));
  };

  // Logout user
  const handleSignOut = async () => {
    try {
      // Panggil endpoint logout
      const response = await fetch("http://localhost:8080/logout", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
        },
        credentials: "include", // Sertakan cookie dalam permintaan
      });

      if (response.ok) {
        window.location.href = "/login";
      } else {
        const errorData = await response.json();
        alert("Logout gagal:", errorData);
      }
    } catch (error) {
      alert("Terjadi kesalahan saat melakukan logout:", error);
    }
  };

  // Fetch events dan set waktu event
  useEffect(() => {
    const fetchNotifications = () => {
        GetNotification((data) => {
            const groupedNotifications = {
                JadwalCuti: [],
                JadwalRapat: [],
                BookingRapat: [],
                TimelineWallpaperDesktop: [],
                TimelineProject: [],
            };
            data.forEach((event) => {
                if (groupedNotifications[event.category]) {
                    groupedNotifications[event.category].push(event);
                }
                // Misalnya, set waktu event untuk event pertama
                if (event.category === 'JadwalRapat') {
                    setEventTime(new Date(event.start)); // Asumsi 'start' adalah waktu mulai event
                }
            });
            setNotification(groupedNotifications);
        });
    };

    fetchNotifications();
    const intervalId = setInterval(fetchNotifications, 60000); // Refresh setiap 1 menit

    return () => clearInterval(intervalId);
  }, []);

  // Hapus Notif
  const handleDelete = async (id) => {
    Swal.fire({
      title: "Apakah Anda yakin?",
      text: "Anda akan menghapus Notif ini!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Ya, saya yakin",
      cancelButtonText: "Batal",
    }).then(async (result) => {
      if (result.isConfirmed) {
        try {
          await deleteNotification(id); // hapus data di API
          setNotification((prevData) => {
            // Pastikan prevData adalah array
            const updatedNotifications = { ...prevData }; // Salin objek sebelumnya
            Object.keys(updatedNotifications).forEach((category) => {
              updatedNotifications[category] = updatedNotifications[
                category
              ].filter((event) => event.id !== id);
            });
            return updatedNotifications; // Kembalikan objek yang diperbarui
          });
        } catch (error) {
          Swal.fire("Gagal!", "Error saat hapus Notif:", error);
        }
      }
    });
  };

  // Fungsi untuk menghandle perubahan pada checkbox
  const handleSelectNotification = (id) => {
    setSelectedNotifications((prevSelected) =>
      prevSelected.includes(id)
        ? prevSelected.filter((selectedId) => selectedId !== id)
        : [...prevSelected, id]
    );
  };

  // Fungsi untuk menghandle select all
  const handleSelectAll = (category) => {
    const allIds = notification[category].map((event) => event.id);
    setSelectedNotifications((prevSelected) =>
      allIds.every((id) => prevSelected.includes(id))
        ? prevSelected.filter((id) => !allIds.includes(id))
        : [...prevSelected, ...allIds.filter((id) => !prevSelected.includes(id))] 
    );
  };

  // Fungsi untuk menghapus notifikasi yang dipilih
  const handleDeleteSelected = async () => {
    try {
      for (const id of selectedNotifications) {
        await deleteNotification(id);
      }
      setNotification((prevData) => {
        const updatedNotifications = { ...prevData };
        Object.keys(updatedNotifications).forEach((category) => {
          updatedNotifications[category] = updatedNotifications[category].filter(
            (event) => !selectedNotifications.includes(event.id)
          );
        });
        return updatedNotifications;
      });
      setSelectedNotifications([]);
    } catch (error) {
      Swal.fire("Gagal!", "Error saat hapus Notif:", error);
    }
  };

  return (
    <div className="grid grid-cols-2fr">
      <Sidebar
        img="../../../public/images/logobjb.png"
        title="IT Security"
        username={userDetails.username}
        email={userDetails.email}
      >
        <SidebarItem
          href="/dashboard"
          text="Dashboard"
          icon={<MdOutlineDashboard />}
        />
        <SidebarCollapse
          text="Dokumen"
          icon={<HiOutlineClipboardDocumentList />}
        >
          <SidebarItem href="/memo" text="Memo" />
          <SidebarItem href={"/berita-acara"} text="Berita Acara" />
          <SidebarItem href="/surat" text="Surat" />
          <SidebarItem href="/sk" text="Sk" />
          <SidebarItem href="/perjalanan-dinas" text="Perjalanan Dinas" />
        </SidebarCollapse>
        <SidebarCollapse text="Project" icon={<GoProjectSymlink />}>
          <SidebarItem href="/project" text="Project" />
          <SidebarItem href="/base-project" text="Base Project" />
        </SidebarCollapse>
        <SidebarCollapse text="Kegiatan" icon={<GrPlan />}>
          <SidebarItem
            href="/timeline-desktop"
            text="Timeline Wallpaper Desktop"
          />
          <SidebarItem href="/booking-rapat" text="Booking Ruang Rapat" />
          <SidebarItem href="/jadwal-rapat" text="Jadwal Rapat" />
          <SidebarItem href="/jadwal-cuti" text="Jadwal Cuti" />
          <SidebarItem href="/meeting" text="Meeting" />
        </SidebarCollapse>
        <SidebarCollapse text="Weekly Timeline" icon={<GrPlan />}>
          <SidebarItem href="/timeline-project" text="Timeline Project" />
          <SidebarItem href="/meeting-schedule" text="Meeting Schedule" />
        </SidebarCollapse>
        <SidebarCollapse text="Informasi" icon={<SlEnvolopeLetter />}>
          <SidebarItem href="/surat-masuk" text="Surat Masuk" />
          <SidebarItem href="/surat-keluar" text="Surat Keluar" />
          <SidebarItem href="/arsip" text="Arsip" />
        </SidebarCollapse>
        {userDetails.role !== "user" && (
          <SidebarItem href="/user" text="User" icon={<FiUsers />} />

        )}

        {userDetails.role !== "user" && (
          <SidebarItem
            href="/request"
            text="Request"
            icon={<VscGitPullRequestGoToChanges />}
          />
        )}
        <SidebarItem
          onClick={handleSignOut}
          text="Logout"
          icon={<BiLogOut />}
        />
      </Sidebar>
      <div className="grid grid-rows-2fr h-screen">
        <header className="mx-4 mt-2 flex justify-between border-b-2 border-gray-100">
          <div className="flex gap-2 items-end m-2">
            <div>
              <Label className="block text-sm">Halaman</Label>
              <Label className="block truncate text-sm font-medium ">
                <b className="uppercase">{services}</b>
              </Label>
            </div>
          </div>
          <div className="flex items-center gap-4 m-2">
            <Label className="truncate text-sm font-medium ring-2 p-1.5 rounded bg-slate-50">
              {RealtimeClock()}
            </Label>
            <Label className="truncate text-sm font-medium ring-2 p-1.5 rounded bg-slate-50">
              {RealtimeDate()}
            </Label>
            <Dropdown
              arrowIcon={false}
              inline
              label={
                <div className="relative flex items-center">
                  {notification &&
                    Object.values(notification).some(
                      (category) => category.length > 0
                    ) && (
                      <div className="absolute -translate-x-[3px] rounded-full bg-green-400">
                        <div className="w-full text-xs text-white px-[5px]">
                          {Object.values(notification).reduce(
                            (total, category) => total + category.length,
                            0
                          )}
                        </div>
                      </div>
                    )}
                  <svg
                    className="w-[34px] h-[34px] text-slate-800 dark:text-white"
                    aria-hidden="true"
                    xmlns="http://www.w3.org/2000/svg"
                    width="24"
                    height="24"
                    fill="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      d="M17.133 12.632v-1.8a5.406 5.406 0 0 1-4.154-5.262.955.955 0 0 0 .021-.106V3.1a1 1 0 0 0-2 0v2.364a.955.955 0 0 0 .021.106 5.406 5.406 0 0 0-4.154 5.262v1.8C6.867 15.018 5 15.614 5 16.807 5 17.4 5 18 5.538 18h12.924C19 18 19 17.4 19 16.807c0-1.193-1.867-1.789-1.867-4.175ZM10 6h4V4h-4v2Zm1 4a1 1 0 1 0-2 0v8a1 1 0 1 0 2 0v-8Zm4 0a1 1 0 1 0-2 0v8a1 1 0 1 0 2 0v-8Z"
                      clipRule="evenodd"
                    />
                  </svg>
                </div>
              }
            >
              <Dropdown.Header>
                <h1 className="text-base">Notification</h1>
                <button onClick={handleDeleteSelected} className="text-red-600">
                  Hapus yang Dipilih
                </button>
              </Dropdown.Header>
              <div className="p-2 grid grid-cols-2 gap-2">
                {Object.keys(filter).map((category) => (
                  <div key={category} className="flex items-center">
                    <input
                      type="checkbox"
                      checked={filter[category]}
                      onChange={() => handleFilterChange(category)}
                      className="mr-2"
                    />
                    <label className="text-sm">{category}</label>
                  </div>
                ))}
              </div>
              <div className="max-h-[50vh] overflow-auto scrollbar-thin scrollbar-thumb-gray-400 scrollbar-track-gray-200">
                {Object.keys(notification).every(
                  (category) => notification[category].length === 0
                ) && (
                    <span className="text-sm text-gray-600">
                      <Badge color="warning" className="m-3">
                        Tidak ada notifikasi
                      </Badge>
                    </span>
                  )}
                {Object.keys(notification).map((category) => (
                  <div key={category}>
                    {filter[category] && notification[category].length > 0 && (
                      <div>
                        <div className="flex items-center ms-[1rem]">
                          <input
                            type="checkbox"
                            onChange={() => handleSelectAll(category)}
                            checked={notification[category].every((event) =>
                              selectedNotifications.includes(event.id)
                            )}
                            className="mr-2"
                          />
                          <h2 className="m-2 font-bold text-xl">{category}</h2>
                        </div>
                        <div className="max-h-[50vh] overflow-auto scrollbar-thin scrollbar-thumb-gray-400 scrollbar-track-gray-200">
                          {notification[category].map((event) => {
                            const formattedStart = format(
                              event.start,
                              "dd MMMM HH:mm",
                              {
                                locale: idLocale,
                              }
                            );
                            return (
                              <Dropdown.Item key={event.id} className="flex w-full justify-between">
                                <div className="flex items-center gap-">
                                  <input
                                    type="checkbox"
                                    checked={selectedNotifications.includes(event.id)}
                                    onChange={() => handleSelectNotification(event.id)}
                                  />
                                  <div className="grid grid-cols-1 grid-rows-2">
                                    <span className="text-start ms-2 font-bold text-base truncate w-48">
                                      {event.title}
                                    </span>
                                    <span>
                                      Pada Waktu {formattedStart}
                                    </span>
                                  </div>
                                </div>
                                <div
                                  className="block text-sm truncate cursor-pointer hover:scale-110 text-red-600 rounded transition-all"
                                  onClick={() => {
                                    handleDelete(event.id);
                                  }}
                                >
                                  <svg
                                    className="w-6 h-6"
                                    aria-hidden="true"
                                    xmlns="http://www.w3.org/2000/svg"
                                    width="24"
                                    height="24"
                                    fill="currentColor"
                                    viewBox="0 0 24 24"
                                  >
                                    <path
                                      fillRule="evenodd"
                                      d="M8.586 2.586A2 2 0 0 1 10 2h4a2 2 0 0 1 2 2v2h3a1 1 0 1 1 0 2v12a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V8a1 1 0 0 1 0-2h3V4a2 2 0 0 1 .586-1.414ZM10 6h4V4h-4v2Zm1 4a1 1 0 1 0-2 0v8a1 1 0 1 0 2 0v-8Zm4 0a1 1 0 1 0-2 0v8a1 1 0 1 0 2 0v-8Z"
                                      clipRule="evenodd"
                                    />
                                  </svg>
                                </div>
                              </Dropdown.Item>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </Dropdown>
          </div>
        </header>
        <div className="mt-4 px-2 w-full overflow-auto">{children}</div>
      </div>
    </div>
  );
};
export default App;
