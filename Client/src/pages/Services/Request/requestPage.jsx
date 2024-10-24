import React, { useState, useEffect } from "react";
import axios from "axios";
import App from "../../../components/Layouts/App";
import { jwtDecode } from "jwt-decode";
import Swal from "sweetalert2";
import { useToken } from "../../../context/TokenContext";

export const RequestPage = () => {
    const [formConfig, setFormConfig] = useState({
        fields: [
            // ... existing fields ...
        ],
        services: "Request",
    });

    const [conflictRequests, setConflictRequests] = useState([]);

    const { token } = useToken();

    let userRole = "";
    if (token) {
        const decoded = jwtDecode(token);
        userRole = decoded.role;
    }

    useEffect(() => {
        if (token) {
            fetch("http://localhost:8080/request?status=pending", { // Tambahkan query parameter status=pending
                method: "GET",
                headers: {
                    Authorization: `Bearer ${token}`,
                },
                credentials: "include",
            })
                .then(response => {
                    if (!response.ok) {
                        throw new Error('Network response was not ok');
                    }
                    return response.json();
                })
                .then(data => {
                    console.log("Data received from server:", data); // Tambahkan log ini
                    if (data && data.request) {
                        setConflictRequests(data.request);
                    } else {
                        console.error("Data format is incorrect or 'booking' is missing");
                    }
                })
                .catch(error => console.error("Error fetching data:", error));
        }
    }, [token]);

    const handleAccept = async (id) => { // Tambahkan parameter id
        const result = await Swal.fire({
            icon: "success",
            title: "Acc Request",
            text: "Anda akan acc request ini?",
            showCancelButton: true,
            confirmButtonText: "Ya, saya yakin",
            cancelButtonText: "Batal",
        });
        if (result.isConfirmed) {
            try {
                await axios.get(`http://localhost:8080/AccRequest/${id}`); // Gunakan id dalam URL
                Swal.fire({
                    icon: "success",
                    title: "Berhasil!",
                    text: "Data berhasil diupdate ke Excel",
                    showConfirmButton: false,
                    timer: 1500,
                });
                window.location.reload();
            } catch (error) {
                Swal.fire("Gagal!", "Error saat update data:", "error");
            }
        }
    };

    const handleCancel = async (id) => {
        const result = await Swal.fire({
            icon: "error",
            title: "Cancel Request",
            text: "Anda akan mengcancel request ini?",
            showCancelButton: true,
            confirmButtonText: "Ya, saya yakin",
            cancelButtonText: "Batal",
        });
        if (result.isConfirmed) {
            try {
                await axios.get(`http://localhost:8080/CancelRequest/${id}`);
                Swal.fire({
                    icon: "success",
                    title: "Berhasil!",
                    text: "Data berhasil di cancel",
                    showConfirmButton: false,
                    timer: 1500,
                });
                window.location.reload();
            } catch (error) {
                Swal.fire("Gagal!", "Error saat cancel data:" + error, "error");
            }

        }
    };

    return (
        <App services={formConfig.services}>
            <div style={{ display: "flex", flexWrap: "wrap" }}>
                {Array.isArray(conflictRequests) && conflictRequests.length > 0 ? (
                    conflictRequests.map((conflict, index) => {
                        console.log("Rendering conflict:", conflict); // Tambahkan log ini
                        return (
                            <div key={index} style={{
                                border: "1px solid #ccc",
                                padding: "20px",
                                borderRadius: "8px",
                                marginBottom: "20px",
                                width: "calc(50% - 20px)", // Setengah layar dengan margin
                                marginRight: index % 2 === 0 ? "20px" : "0" // Tambahkan margin kanan untuk card di sebelah kiri
                            }}>
                                <h3>Data Bentrok</h3>
                                <p>
                                    Jadwal rapat <strong>{conflict.title}</strong> bentrok pada tanggal <strong>{conflict.start}</strong>
                                </p>
                                <button
                                    onClick={() => handleAccept(conflict.id)} // Pass ID to handleAccept
                                    style={{ marginRight: "10px" }}
                                    className="text-white bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm w-auto px-5 py-2.5 text-center mt-2"
                                >
                                    Acc
                                </button>
                                <button
                                    onClick={() => handleCancel(conflict.id)} // Pass ID to handleCancel
                                    className="text-white bg-red-700 hover:bg-red-800 focus:ring-4 focus:outline-none focus:ring-red-300 font-medium rounded-lg text-sm w-auto px-5 py-2.5 text-center mt-2"
                                >
                                    Cancel
                                </button>
                            </div>
                        );
                    })
                ) : (
                    <div style={{ width: "100%",height: "80vh", display: "flex", justifyContent: "center", alignItems: "center", marginTop: "20px" }}>
                        Tidak ada jadwal bentrok
                    </div>
                )}
            </div>
        </App>
    );
};
