import React, { useState, useRef } from "react";
import logo from "./knack_RCM_text.png";
import { ToastContainer, toast } from "react-toastify";
import "./App.css";
import "react-toastify/dist/ReactToastify.css";
import { useDropzone } from "react-dropzone";
import {
  Box,
  Typography,
  Button,
  AppBar,
  Toolbar,
  Paper,
  CircularProgress,
  Grid,
} from "@mui/material";
import { TaskAlt, HourglassEmpty, PlayArrow, Stop, Download, Replay } from "@mui/icons-material";
import * as XLSX from "xlsx";

// --- CONFIGURATION ---
// Helper function to safely get the API URL regardless of the build tool (Vite or Webpack/CRA)
const getApiEndpoint = () => {
  try {
    // Check for Vite environment
    // @ts-ignore
    if (typeof import.meta !== "undefined" && import.meta.env && import.meta.env.VITE_API_URL) {
      // @ts-ignore
      return import.meta.env.VITE_API_URL;
    }
    // Check for Create React App / Webpack environment
    if (typeof process !== "undefined" && process.env && process.env.REACT_APP_API_URL) {
      return process.env.REACT_APP_API_URL;
    }
  } catch (e) {
    // Ignore errors and fall back to default
  }
  // Default local development URL
  return "http://127.0.0.1:5000/extract";
};

const API_ENDPOINT = getApiEndpoint();

// Limit file size to 25MB to prevent server timeouts/crashes
const MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024; 

export default function App() {
  const [files, setFiles] = useState([]);
  const [extractedRows, setExtractedRows] = useState([]);
  const [progress, setProgress] = useState({ current: 0, total: 0, running: false });

  // We use a Ref for stopping because it updates immediately inside the loop
  const stopRef = useRef(false);
  const controllerRef = useRef(null);

  // Helper to check if we have any stopped files to trigger "Resume" UI
  const hasStoppedFiles = files.some((f) => f.status === "stopped" || f.status === "pending");
  const isAllDone = files.length > 0 && files.every((f) => f.status === "done");

  const { getRootProps, getInputProps } = useDropzone({
    accept: { "application/pdf": [] },
    onDrop: (acceptedFiles) => {
      if (acceptedFiles.length === 0) {
        // If dropzone rejects files (e.g. non-pdf), this might trigger.
        // Ideally we check fileRejections from useDropzone hook, but this is a basic guard.
        return;
      }
      // Only allow adding files if we aren't currently running
      if (progress.running) {
        toast.warning("Cannot add files while processing");
        return;
      }

      // Validate files
      const validFiles = [];
      acceptedFiles.forEach((file) => {
        if (file.size > MAX_FILE_SIZE_BYTES) {
          toast.error(`Skipped ${file.name}: File too large (>25MB)`);
        } else {
          validFiles.push({
            file: file,
            status: "pending",
          });
        }
      });

      if (validFiles.length > 0) {
        setFiles((prev) => [...prev, ...validFiles]);
        toast.success(`${validFiles.length} PDF(s) added`);
      }
    },
    onDropRejected: () => {
        toast.warning("Only PDF files are allowed.");
    }
  });

  // FUNCTION: Download the Combined Excel Data
  const downloadCombinedExcel = () => {
    if (extractedRows.length === 0) {
      toast.warning("No extracted data to download", { autoClose: 1000 });
      return;
    }

    const ws = XLSX.utils.json_to_sheet(extractedRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "All Data");

    // Get today's date in system locale format and sanitize slashes
    const dateStr = new Date().toLocaleDateString().replace(/\//g, "-");
    const fileName = `report_${dateStr}.xlsx`;

    // This creates one file named like "report_11-20-2025.xlsx" (depending on locale)
    XLSX.writeFile(wb, fileName);
    toast.success(`${fileName} Downloaded! 📥`, { autoClose: 1500 });
  };

  // STOP EXTRACTION
  const stopExtraction = () => {
    stopRef.current = true; // Set the flag immediately
    if (controllerRef.current) {
      controllerRef.current.abort(); // Kill the current fetch request
    }
    toast.warn("Stopping process... 🛑", { autoClose: 1000 });
  };

  // HANDLE EXTRACT (Acts as Start AND Resume)
  const handleExtract = async () => {
    if (files.length === 0) {
      toast.error("Please select PDF folder first", { autoClose: 1000 });
      return;
    }

    try {
      stopRef.current = false; // Reset stop flag
      setProgress({ current: 0, total: files.length, running: true });

      // If we are starting fresh (extractedRows is empty), toast starting. 
      // If resuming, toast resuming.
      if (extractedRows.length === 0 && !hasStoppedFiles) {
        toast.info("Starting extraction...", { autoClose: 1000 });
      } else {
        toast.info("Resuming extraction...", { autoClose: 1000 });
      }

      let currentBatchData = [];

      for (let i = 0; i < files.length; i++) {
        // ------------------------------------------
        // 1. SKIP COMPLETED FILES (Resume Logic)
        // ------------------------------------------
        if (files[i].status === "done") {
          // Just update progress index visually
          setProgress((prev) => ({ ...prev, current: i + 1 }));
          continue;
        }

        // ------------------------------------------
        // 2. CHECK STOP FLAG
        // ------------------------------------------
        if (stopRef.current) {
          // Mark the current file (and subsequent ones) as stopped
          setFiles((prev) =>
            prev.map((f, idx) =>
              idx >= i && f.status !== "done" ? { ...f, status: "stopped" } : f
            )
          );
          break; // Exit the loop
        }

        const fileObj = files[i];

        // Update status to processing
        setFiles((prev) =>
          prev.map((f, idx) =>
            idx === i ? { ...f, status: "processing" } : f
          )
        );

        const formData = new FormData();
        formData.append("pdf_file", fileObj.file);

        controllerRef.current = new AbortController();

        try {
          const response = await fetch(API_ENDPOINT, {
            method: "POST",
            body: formData,
            signal: controllerRef.current.signal,
          });

          const result = await response.json();

          if (!response.ok || result.status !== "success") {
            const msg = result.message || `Error ${response.status}`;
            // Log error internally but show user-friendly message
            console.error(`Extraction error for ${fileObj.file.name}:`, msg);
            toast.error(`${fileObj.file.name}: Extraction failed`);
            
            setFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: "error" } : f));
            // We continue to next file even on error, unless you want to stop on error
            continue;
          }

          // ✅ SUCCESS
          if (result.pages && Array.isArray(result.pages)) {
            const fileRows = result.pages.map(p => ({
              FileName: fileObj.file.name,
              Page: p.page,
              ...p.data
            }));

            // Add to local batch
            currentBatchData.push(...fileRows);

            // Also immediately add to main state (safe way)
            setExtractedRows(prev => [...prev, ...fileRows]);
          }

          setFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: "done" } : f));

        } catch (err) {
          // ------------------------------------------
          // 3. HANDLE ABORT (STOPPED DURING FETCH)
          // ------------------------------------------
          if (err.name === "AbortError" || stopRef.current) {
            setFiles((prev) =>
              prev.map((f, idx) =>
                idx === i ? { ...f, status: "stopped" } : f
              )
            );
            // Silent log
            // console.log("Fetch aborted by user");
            break; // Stop loop
          } else {
            console.error("Network/Code Error:", err);
            setFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: "error" } : f));
            toast.error(`${fileObj.file.name}: Network or Server Error`);
          }
        }

        setProgress((prev) => ({ ...prev, current: i + 1 }));
      }

      setProgress({ current: 0, total: 0, running: false });

      // Final check: If we finished and no stop was requested
      if (!stopRef.current) {
        toast.success("All processes finished! ✅", { autoClose: 1500 });
      } else {
        toast.warn("Process paused. Click Resume to continue.", { autoClose: 1000 });
      }

    } catch (err) {
      console.error(err);
      setProgress({ current: 0, total: 0, running: false });
      toast.error("Critical system error occurred", { autoClose: 2500 });
    }
  };

  return (
    <>
      <ToastContainer />
      <AppBar
        position="fixed"
        sx={{
          bgcolor: "#e9eff7ff",
          boxShadow: "0px 4px 12px rgba(0,0,0,0.15)",
        }}
      >
        <Toolbar sx={{ display: "flex", justifyContent: "space-between" }}>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <img src={logo} alt="logo" style={{ height: 45 }} />
          </Box>
          <Typography variant="h5" sx={{ color: "#494d74ff", fontWeight: 700 }}>
            EOB Extracting
          </Typography>
        </Toolbar>
      </AppBar>

      <Grid container sx={{ mt: 10 }}>
        {/* LEFT PANEL */}
        <Grid size={3} sx={{ p: 3, minHeight: "80vh" }}>
          <Box
            {...getRootProps()}
            sx={{
              border: "2px dashed #8FABD4",
              bgcolor: "#F7F9FC",
              p: 5,
              m: 2,
              borderRadius: 2,
              textAlign: "center",
              cursor: progress.running ? "not-allowed" : "pointer",
              opacity: progress.running ? 0.6 : 1,
              "&:hover": { bgcolor: "#eaeff7ff" },
            }}
          >
            <input
              {...getInputProps({
                webkitdirectory: "true",
                directory: "true",
                disabled: progress.running
              })}
            />
            <Typography sx={{ fontWeight: 600, color: "#0A2540" }}>
              Drop Folder OR click to select folder
            </Typography>
            <Typography sx={{ fontSize: 13, color: "#6a0404ff" }}>
               *Folder must contain only PDF files
            </Typography>
          </Box>

          <Grid container spacing={2} justifyContent="center">
            <Grid item xs={12} sx={{ display: "flex", justifyContent: "center", gap: 2 }}>

              {/* BUTTON LOGIC */}
              {progress.running ? (
                <Button
                  variant="contained"
                  color="error"
                  onClick={stopExtraction}
                  sx={{
                    fontWeight: 700,
                    px: 3,
                    gap: 1,
                    "&:hover": { bgcolor: "#b71c1c", transform: "scale(1.05)" },
                  }}
                >
                  Stop Processing
                  <Stop />
                </Button>
              ) : (
                // If not running
                hasStoppedFiles && extractedRows.length > 0 ? (
                  <Button
                    variant="contained"
                    onClick={handleExtract}
                    sx={{
                      bgcolor: "#ff9800", // Orange for Resume
                      fontWeight: 700,
                      px: 3,
                      gap: 1,
                      "&:hover": { bgcolor: "#f57c00", transform: "scale(1.05)" },
                    }}
                  >
                    Resume Extraction
                    <Replay />
                  </Button>
                ) : (
                  <Button
                    variant="contained"
                    onClick={handleExtract}
                    disabled={files.length === 0 || isAllDone}
                    sx={{
                      bgcolor: "#2e7d32",
                      fontWeight: 700,
                      px: 3,
                      gap: 1,
                      "&:hover": { bgcolor: "#1b5e20", transform: "scale(1.05)" },
                    }}
                  >
                    {isAllDone ? "All Done" : "Extract Now"}
                    <PlayArrow />
                  </Button>
                )
              )}
            </Grid>

            {/* --- DOWNLOAD BUTTON --- */}
            {extractedRows.length > 0 && !progress.running && (
              <Grid item xs={12} sx={{ display: "flex", justifyContent: "center", mt: 3 }}>
                <Button
                  variant="contained"
                  onClick={downloadCombinedExcel}
                  sx={{
                    bgcolor: "#053b56ff",
                    fontWeight: 700,
                    px: 2,
                    py: 1,
                    width: "100%",
                    transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
                    "&:hover": {
                      bgcolor: "#02350cff",
                      transform: "scale(1.05)",
                    },
                  }}

                >
                  Download Report <Download sx={{ mr: 1 }} />
                </Button>
              </Grid>
            )}
          </Grid>
        </Grid>

        {/* RIGHT PANEL */}
        <Grid size={9} sx={{ p: 3, minHeight: "80vh" }}>
          {files.length > 0 && (
            <Paper elevation={3} sx={{ p: 3, borderRadius: 2, bgcolor: "white" }}>
              <Typography
                sx={{ mb: 2, fontWeight: 700, color: "#0A2540", fontSize: "1.1rem" }}
              >
                Uploaded PDF Files : {files.length}
              </Typography>

              <Box sx={{ maxHeight: 425, overflowY: "auto", borderRadius: 1 }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead
                    style={{
                      position: "sticky",
                      top: 0,
                      backgroundColor: "#053b56ff",
                      color: "#fff",
                      zIndex: 1,
                    }}
                  >
                    <tr>
                      <th style={{ padding: "10px", textAlign: "left" }}>File Name</th>
                      <th style={{ padding: "10px", textAlign: "left" }}>Size (KB)</th>
                      <th style={{ padding: "10px", textAlign: "center" }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {files.map((f, i) => (
                      <tr
                        key={i}
                        style={{
                          borderBottom: "1px solid #e0e0e0",
                          backgroundColor:
                            f.status === "processing"
                              ? "#e3f2fd"
                              : f.status === "error"
                                ? "#ffebee"
                                : f.status === "stopped"
                                  ? "#fff3e0" // Light orange for stopped
                                  : f.status === "done"
                                    ? "#e8f5e9" // Light green for done
                                    : "transparent",
                        }}
                      >
                        <td style={{ padding: "10px" }}>{f.file.name}</td>
                        <td style={{ padding: "10px" }}>
                          {(f.file.size / 1024).toFixed(2)}
                        </td>
                        <td style={{ padding: "10px", textAlign: "center" }}>
                          {f.status === "pending" && <HourglassEmpty sx={{ color: "#FFB300" }} />}
                          {f.status === "processing" && <CircularProgress size={20} />}
                          {f.status === "done" && <TaskAlt sx={{ color: "#2e7d32" }} />}
                          {f.status === "error" && (
                            <Typography sx={{ color: "#d32f2f", fontWeight: 600 }}>❌</Typography>
                          )}
                          {f.status === "stopped" && (
                            <Typography sx={{ color: "#ff9800", fontWeight: 600 }}>⏸️ Paused</Typography>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Box>
            </Paper>
          )}

          {progress.running && (
            <Box
              sx={{
                position: "fixed",
                bottom: 20,
                right: 20,
                color: "#ecedf1ff",
                bgcolor: "#0b6b02ff",
                p: 2,
                borderRadius: 2,
                boxShadow: "0px 4px 12px rgba(0,0,0,0.15)",
                fontWeight: 600,
                display: "flex",
                alignItems: "center",
                gap: 1,
                zIndex: 1000,
              }}
            >
              <HourglassEmpty sx={{ animation: "spin 1s linear infinite" }} />
              Processing File {progress.current + 1} of {progress.total}...
            </Box>
          )}
        </Grid>
      </Grid>
    </>
  );
}