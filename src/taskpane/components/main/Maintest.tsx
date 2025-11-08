// // // src/taskpane/components/main/Home.tsx (updated insertIntoWord function)
// // import React, { useEffect, useState } from "react";
// // import {
// //   Box,
// //   Typography,
// //   Divider,
// //   Button,
// //   Modal,
// //   Paper,
// //   Stack,
// //   IconButton,
// //   Card,
// //   CardContent,
// //   CardActions,
// //   CircularProgress,
// //   Snackbar,
// //   Alert,
// // } from "@mui/material";
// // import CloseIcon from "@mui/icons-material/Close";
// // import AppHeader from "./AppHeader";
// // import { fetchAddonReports, Inspection } from "../services/AddonReports";
// // import stringSimilarity from "string-similarity";

// // export default function Home() {
// //   const [inspections, setInspections] = useState<Inspection[]>([]);
// //   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
// //   const [loading, setLoading] = useState(false);
// //   const [error, setError] = useState<string | null>(null);
// //   const [message, setMessage] = useState<string | null>(null);

// //   // ---------------------------
// //   // Fetch Addon Reports
// //   // ---------------------------
// //   const getReports = async () => {
// //     setLoading(true);
// //     setError(null);
// //     try {
// //       const data = await fetchAddonReports();
// //       console.log("Fetched inspections:", data);
// //       setInspections(data);
// //       setMessage("Inspection reports loaded successfully!");
// //     } catch (err: any) {
// //       console.error(err);
// //       setError("Failed to fetch inspection reports. Please try again.");
// //     } finally {
// //       setLoading(false);
// //     }
// //   };

// //   useEffect(() => {
// //     getReports();
// //   }, []);

// //   // ---------------------------
// //   // Client-specific section titles
// //   // ---------------------------
// //   const templatesByClient: Record<string, string[]> = {
// //     WoodTech: [
// //       "Introduction / Scope",
// //       "Purpose of Inspection",
// //       "Reference Documents and Standards",
// //       "Findings",
// //       "Photos",
// //       "Recommendations", // Added for more flexibility
// //       "Summary" // Added for more flexibility
// //     ],
// //     SteelWorks: [
// //       "Inspection Overview",
// //       "Observations",
// //       "Defects",
// //       "Conclusion",
// //       "Report Details", // Added for more flexibility
// //     ],
// //   };

// //   // ---------------------------
// //   // Insert into Word
// //   // ---------------------------
// //   const insertIntoWord = async (inspection: Inspection) => {
// //     const clientName = inspection.attr?.client || "WoodTech";
// //     const clientSections = templatesByClient[clientName] || templatesByClient["WoodTech"];

// //     // Define the data points from your inspection object and their corresponding default report section names
// //     const dataToInsert = [
// //       {
// //         field: "reportNumber",
// //         label: "Report #",
// //         value: inspection.attr?.no || "N/A",
// //         defaultSection: "Report Details", // A general section for basic info
// //         insertAsParagraph: true, // Insert as a standalone paragraph
// //       },
// //       {
// //         field: "location",
// //         label: "Location",
// //         value: inspection.attr?.location || "Unknown",
// //         defaultSection: "Report Details",
// //         insertAsParagraph: true,
// //       },
// //       {
// //         field: "material",
// //         label: "Material",
// //         value: inspection.attr?.material || "Unknown Material",
// //         defaultSection: "Report Details",
// //         insertAsParagraph: true,
// //       },
// //       {
// //         field: "inspector",
// //         label: "Inspector",
// //         value: inspection.attr?.inspector || "No Inspector",
// //         defaultSection: "Report Details",
// //         insertAsParagraph: true,
// //       },
// //       {
// //         field: "date",
// //         label: "Date",
// //         value: inspection.attr?.date
// //           ? new Date(inspection.attr.date).toLocaleDateString()
// //           : "Unknown",
// //         defaultSection: "Report Details",
// //         insertAsParagraph: true,
// //       },
// //       {
// //         field: "observations",
// //         label: "Observations",
// //         value:
// //           inspection.ai_notes?.inspectorObservations ||
// //           inspection.notes?.inspectorObservations ||
// //           "No observations available.",
// //         defaultSection: "Findings", // Primary place for observations
// //         insertAsParagraph: false, // Will be inserted after the section title
// //       },
// //       // You can add more fields here if needed, e.g., for recommendations or summary if they exist in ai_notes
// //       {
// //         field: "summary",
// //         label: "Summary",
// //         value: inspection.ai_notes?.inspectorObservations?.match(/\*\*Summary\*\*\s*([\s\S]*)/)?.[1]?.trim() || "No summary available.",
// //         defaultSection: "Summary",
// //         insertAsParagraph: false,
// //       },
// //       {
// //         field: "nonConformances",
// //         label: "Non-Conformances",
// //         value: inspection.ai_notes?.inspectorObservations?.match(/\*\*Non-Conformances\*\*\s*([\s\S]*?)(?=\*\*Summary\*\*|$)/)?.[1]?.trim() || "No non-conformances noted.",
// //         defaultSection: "Defects", // Or a custom section like "Non-Conformances"
// //         insertAsParagraph: false,
// //       }
// //     ];

// //     try {
// //       await Word.run(async (context) => {
// //         const paragraphs = context.document.body.paragraphs;
// //         paragraphs.load("items");
// //         await context.sync();

// //         for (const dataItem of dataToInsert) {
// //           const { field, label, value, defaultSection, insertAsParagraph } = dataItem;

// //           // Find the best matching section in the client's template
// //           const { bestMatch } = stringSimilarity.findBestMatch(defaultSection, clientSections);
// //           const matchedSectionTitle = bestMatch.target;

// //           // Find the paragraph in the Word document that matches this section title
// //           const sectionParagraph = paragraphs.items.find((p) =>
// //             p.text.toLowerCase().includes(matchedSectionTitle.toLowerCase())
// //           );

// //           if (sectionParagraph) {
// //             if (insertAsParagraph) {
// //               // For simple key-value pairs (like Location, Material)
// //               sectionParagraph.insertParagraph(`${label}: ${value}`, "After");
// //             } else {
// //               // For longer text blocks (like Observations, Summary)
// //               if (value && value !== "No observations available." && value !== "No summary available." && value !== "No non-conformances noted.") {
// //                  sectionParagraph.insertParagraph(`${label}:\n${value}`, "After");
// //               } else {
// //                  sectionParagraph.insertParagraph(`${label}: ${value}`, "After");
// //               }
// //             }
// //           } else {
// //             console.warn(`Could not find a section matching "${matchedSectionTitle}" for ${label}.`);
// //             // Optionally, insert the data at the end of the document or a generic section
// //             // context.document.body.insertParagraph(`${label}: ${value}`, "End");
// //           }
// //         }

// //         await context.sync();
// //       });

// //       setMessage(`✅ Report inserted successfully for ${clientName}!`);
// //     } catch (error) {
// //       console.error("Error inserting into Word:", error);
// //       setError("⚠️ Failed to insert report into Word.");
// //     }
// //   };

// //   // ---------------------------
// //   // Render UI (remains largely the same, but for context)
// //   // ---------------------------
// //   return (
// //     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
// //       <AppHeader />

// //       {/* Loading Overlay */}
// //       {loading && (
// //         <Box
// //           position="fixed"
// //           top={0}
// //           left={0}
// //           right={0}
// //           bottom={0}
// //           display="flex"
// //           alignItems="center"
// //           justifyContent="center"
// //           bgcolor="rgba(255,255,255,0.7)"
// //           zIndex={9999}
// //         >
// //           <CircularProgress size={60} />
// //         </Box>
// //       )}

// //       {/* Snackbar Alerts */}
// //       <Snackbar
// //         open={!!message || !!error}
// //         autoHideDuration={4000}
// //         onClose={() => {
// //           setMessage(null);
// //           setError(null);
// //         }}
// //         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
// //       >
// //         {message ? (
// //           <Alert severity="success" sx={{ width: "100%" }}>
// //             {message}
// //           </Alert>
// //         ) : error ? (
// //           <Alert severity="error" sx={{ width: "100%" }}>
// //             {error}
// //           </Alert>
// //         ) : null}
// //       </Snackbar>

// //       <Box display="flex" flexGrow={1}>
// //         {/* Sidebar */}
// //         <Box
// //           p={3}
// //           width={340}
// //           bgcolor="#fff"
// //           borderRight="1px solid #ddd"
// //           display="flex"
// //           flexDirection="column"
// //           sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
// //         >
// //           <Typography variant="h6" mb={1} sx={{ color: "#333", fontWeight: 600 }}>
// //             Inspection Reports
// //           </Typography>
// //           <Divider sx={{ mb: 2 }} />

// //           {/* List of Inspections */}
// //           <Box flexGrow={1} overflow="auto" pr={1}>
// //             {inspections.length > 0 ? (
// //               inspections.map((insp) => (
// //                 <Card
// //                   key={insp.id}
// //                   sx={{
// //                     mb: 2,
// //                     borderRadius: 2,
// //                     boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
// //                     transition: "0.3s",
// //                     "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
// //                   }}
// //                 >
// //                   <CardContent>
// //                     <Typography variant="subtitle1" fontWeight={600}>
// //                       #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
// //                     </Typography>
// //                     <Typography variant="body2" color="text.secondary">
// //                       {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
// //                     </Typography>
// //                   </CardContent>

// //                   <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
// //                     <Button
// //                       variant="outlined"
// //                       size="small"
// //                       onClick={() => setSelectedInspection(insp)}
// //                     >
// //                       Preview
// //                     </Button>
// //                     <Button
// //                       variant="contained"
// //                       size="small"
// //                       onClick={() => insertIntoWord(insp)}
// //                     >
// //                       Insert
// //                     </Button>
// //                   </CardActions>
// //                 </Card>
// //               ))
// //             ) : (
// //               !loading && <Typography>No inspections found.</Typography>
// //             )}
// //           </Box>

// //           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
// //             Refresh Data
// //           </Button>
// //         </Box>

// //         {/* Modal for Preview */}
// //         <Modal
// //           open={!!selectedInspection}
// //           onClose={() => setSelectedInspection(null)}
// //           sx={{
// //             display: "flex",
// //             alignItems: "center",
// //             justifyContent: "center",
// //             p: 2,
// //           }}
// //         >
// //           <Paper
// //             sx={{
// //               position: "relative",
// //               width: { xs: "100%", sm: "90%", md: 600 },
// //               maxHeight: "86vh",
// //               overflowY: "auto",
// //               p: 3,
// //               borderRadius: 3,
// //               boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
// //             }}
// //           >
// //             <IconButton
// //               onClick={() => setSelectedInspection(null)}
// //               sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
// //             >
// //               <CloseIcon />
// //             </IconButton>

// //             {selectedInspection && (
// //               <>
// //                 <Typography variant="h6" gutterBottom>
// //                   Report #{selectedInspection.attr?.no}
// //                 </Typography>
// //                 <Divider sx={{ mb: 2 }} />
// //                 <Stack spacing={1}>
// //                   <Typography>
// //                     <b>Location:</b> {selectedInspection.attr?.location}
// //                   </Typography>
// //                   <Typography>
// //                     <b>Material:</b> {selectedInspection.attr?.material}
// //                   </Typography>
// //                   <Typography>
// //                     <b>Inspector:</b> {selectedInspection.attr?.inspector}
// //                   </Typography>
// //                   <Typography>
// //                     <b>Date:</b>{" "}
// //                     {selectedInspection.attr?.date
// //                       ? new Date(selectedInspection.attr.date).toLocaleDateString()
// //                       : "Unknown"}
// //                   </Typography>
// //                   <Typography sx={{ whiteSpace: "pre-wrap" }}>
// //                     <b>Observations:</b>{" "}
// //                     {selectedInspection.ai_notes?.inspectorObservations ||
// //                       selectedInspection.notes?.inspectorObservations ||
// //                       "No observations."}
// //                   </Typography>
// //                 </Stack>
// //               </>
// //             )}
// //           </Paper>
// //         </Modal>
// //       </Box>
// //     </Box>
// //   );
// // }



// // src/taskpane/components/main/Home.tsx (updated insertIntoWord function)
// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Divider,
//   Button,
//   Modal,
//   Paper,
//   Stack,
//   IconButton,
//   Card,
//   CardContent,
//   CardActions,
//   CircularProgress,
//   Snackbar,
//   Alert,
// } from "@mui/material";
// import CloseIcon from "@mui/icons-material/Close";
// import AppHeader from "./AppHeader";
// import { fetchAddonReports, Inspection } from "../services/AddonReports";
// import stringSimilarity from "string-similarity";

// export default function Home() {
//   const [inspections, setInspections] = useState<Inspection[]>([]);
//   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string | null>(null);
//   const [message, setMessage] = useState<string | null>(null);

//   const getReports = async () => {
//     setLoading(true);
//     setError(null);
//     try {
//       const data = await fetchAddonReports();
//       setInspections(data);
//       setMessage("Inspection reports loaded successfully!");
//     } catch (err: any) {
//       console.error(err);
//       setError("Failed to fetch inspection reports. Please try again.");
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     getReports();
//   }, []);

//   const templatesByClient: Record<string, string[]> = {
//     WoodTech: [
//       "Introduction / Scope",
//       "Purpose of Inspection",
//       "Reference Documents and Standards",
//       "Findings",
//       "Photos",
//       "Recommendations",
//       "Summary"
//     ],
//     SteelWorks: [
//       "Inspection Overview",
//       "Observations",
//       "Defects",
//       "Conclusion",
//       "Report Details",
//     ],
//   };

//   const cleanAndSplitText = (text: string | undefined): string[] => {
//     if (!text) return [];
//     let cleanedText = text.replace(/\*\*(.*?)\*\*/g, '$1').trim();
//     return cleanedText.split(/\n\s*\n/).map(p => p.trim()).filter(p => p.length > 0);
//   };

//   const insertIntoWord = async (inspection: Inspection) => {
//     const clientName = inspection.attr?.client || "WoodTech";
//     const clientSections = templatesByClient[clientName] || templatesByClient["WoodTech"];

//     const fullObservationsText = inspection.ai_notes?.inspectorObservations || inspection.notes?.inspectorObservations || "";

//     const observations = fullObservationsText.match(/\*\*Inspector Observations\*\*\s*([\s\S]*?)(?=\*\*Non-Conformances\*\*|\*\*Summary\*\*|$)/)?.[1]?.trim();
//     const nonConformances = fullObservationsText.match(/\*\*Non-Conformances\*\*\s*([\s\S]*?)(?=\*\*Summary\*\*|$)/)?.[1]?.trim();
//     const summary = fullObservationsText.match(/\*\*Summary\*\*\s*([\s\S]*)/)?.[1]?.trim();

//     const dataToInsert = [
//       {
//         field: "reportNumber",
//         label: "Report #",
//         value: inspection.attr?.no || "N/A",
//         defaultSection: "Report Details",
//       },
//       {
//         field: "location",
//         label: "Location",
//         value: inspection.attr?.location || "Unknown",
//         defaultSection: "Report Details",
//       },
//       {
//         field: "material",
//         label: "Material",
//         value: inspection.attr?.material || "Unknown Material",
//         defaultSection: "Report Details",
//       },
//       {
//         field: "inspector",
//         label: "Inspector",
//         value: inspection.attr?.inspector || "No Inspector",
//         defaultSection: "Report Details",
//       },
//       {
//         field: "date",
//         label: "Date",
//         value: inspection.attr?.date
//           ? new Date(inspection.attr.date).toLocaleDateString()
//           : "Unknown",
//         defaultSection: "Report Details",
//       },
//       {
//         field: "observations",
//         label: "Observations",
//         value: observations || "No observations available.",
//         defaultSection: "Findings",
//         insertMultipleParagraphs: true,
//       },
//       {
//         field: "nonConformances",
//         label: "Non-Conformances",
//         value: nonConformances || "No non-conformances noted.",
//         defaultSection: "Defects",
//         insertMultipleParagraphs: true,
//       },
//       {
//         field: "summary",
//         label: "Summary",
//         value: summary || "No summary available.",
//         defaultSection: "Summary",
//         insertMultipleParagraphs: true,
//       }
//     ];

//     try {
//       await Word.run(async (context) => {
//         const paragraphs = context.document.body.paragraphs;
//         paragraphs.load("items");
//         await context.sync();

//         for (const dataItem of dataToInsert) {
//           const { field, label, value, defaultSection, insertMultipleParagraphs } = dataItem;

//           const { bestMatch } = stringSimilarity.findBestMatch(defaultSection, clientSections);
//           const matchedSectionTitle = bestMatch.target;

//           const sectionParagraph = paragraphs.items.find((p) =>
//             p.text.toLowerCase().includes(matchedSectionTitle.toLowerCase())
//           );

//           if (sectionParagraph) {
//             if (insertMultipleParagraphs) {
//               const contentParagraphs = cleanAndSplitText(value);

//               // Use Office.InsertLocation.after
//               sectionParagraph.insertParagraph(`${label}:`, "After");
//               await context.sync();

//               let lastInsertedParagraph = sectionParagraph;
//               for (const pText of contentParagraphs) {
//                 // Use Office.InsertLocation.after
//                 lastInsertedParagraph = lastInsertedParagraph.insertParagraph(pText, Word.InsertLocation.after);
//                 await context.sync();
//               }

//             } else {
//               // Use Office.InsertLocation.after
//               sectionParagraph.insertParagraph(`${label}: ${value}`, Word.InsertLocation.after);
//               await context.sync();
//             }
//           } else {
//             console.warn(`Could not find a section matching "${matchedSectionTitle}" for ${label}.`);
//           }
//         }

//         await context.sync();
//       });

//       setMessage(`✅ Report inserted successfully for ${clientName}!`);
//     } catch (error) {
//       console.error("Error inserting into Word:", error);
//       setError("⚠️ Failed to insert report into Word.");
//     }
//   };

//   // ---------------------------
//   // Render UI (remains largely the same)
//   // ---------------------------
//   return (
//     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
//       <AppHeader />

//       {/* Loading Overlay */}
//       {loading && (
//         <Box
//           position="fixed"
//           top={0}
//           left={0}
//           right={0}
//           bottom={0}
//           display="flex"
//           alignItems="center"
//           justifyContent="center"
//           bgcolor="rgba(255,255,255,0.7)"
//           zIndex={9999}
//         >
//           <CircularProgress size={60} />
//         </Box>
//       )}

//       {/* Snackbar Alerts */}
//       <Snackbar
//         open={!!message || !!error}
//         autoHideDuration={4000}
//         onClose={() => {
//           setMessage(null);
//           setError(null);
//         }}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         {message ? (
//           <Alert severity="success" sx={{ width: "100%" }}>
//             {message}
//           </Alert>
//         ) : error ? (
//           <Alert severity="error" sx={{ width: "100%" }}>
//             {error}
//           </Alert>
//         ) : null}
//       </Snackbar>

//       <Box display="flex" flexGrow={1}>
//         {/* Sidebar */}
//         <Box
//           p={3}
//           width={340}
//           bgcolor="#fff"
//           borderRight="1px solid #ddd"
//           display="flex"
//           flexDirection="column"
//           sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
//         >
//           <Typography variant="h6" mb={1} sx={{ color: "#333", fontWeight: 600 }}>
//             Inspection Reports
//           </Typography>
//           <Divider sx={{ mb: 2 }} />

//           {/* List of Inspections */}
//           <Box flexGrow={1} overflow="auto" pr={1}>
//             {inspections.length > 0 ? (
//               inspections.map((insp) => (
//                 <Card
//                   key={insp.id}
//                   sx={{
//                     mb: 2,
//                     borderRadius: 2,
//                     boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
//                     transition: "0.3s",
//                     "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
//                   }}
//                 >
//                   <CardContent>
//                     <Typography variant="subtitle1" fontWeight={600}>
//                       #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
//                     </Typography>
//                     <Typography variant="body2" color="text.secondary">
//                       {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
//                     </Typography>
//                   </CardContent>

//                   <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
//                     <Button
//                       variant="outlined"
//                       size="small"
//                       onClick={() => setSelectedInspection(insp)}
//                     >
//                       Preview
//                     </Button>
//                     <Button
//                       variant="contained"
//                       size="small"
//                       onClick={() => insertIntoWord(insp)}
//                     >
//                       Insert
//                     </Button>
//                   </CardActions>
//                 </Card>
//               ))
//             ) : (
//               !loading && <Typography>No inspections found.</Typography>
//             )}
//           </Box>

//           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
//             Refresh Data
//           </Button>
//         </Box>

//         {/* Modal for Preview */}
//         <Modal
//           open={!!selectedInspection}
//           onClose={() => setSelectedInspection(null)}
//           sx={{
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "center",
//             p: 2,
//           }}
//         >
//           <Paper
//             sx={{
//               position: "relative",
//               width: { xs: "100%", sm: "90%", md: 600 },
//               maxHeight: "86vh",
//               overflowY: "auto",
//               p: 3,
//               borderRadius: 3,
//               boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
//             }}
//           >
//             <IconButton
//               onClick={() => setSelectedInspection(null)}
//               sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
//             >
//               <CloseIcon />
//             </IconButton>

//             {selectedInspection && (
//               <>
//                 <Typography variant="h6" gutterBottom>
//                   Report #{selectedInspection.attr?.no}
//                 </Typography>
//                 <Divider sx={{ mb: 2 }} />
//                 <Stack spacing={1}>
//                   <Typography>
//                     <b>Location:</b> {selectedInspection.attr?.location}
//                   </Typography>
//                   <Typography>
//                     <b>Material:</b> {selectedInspection.attr?.material}
//                   </Typography>
//                   <Typography>
//                     <b>Inspector:</b> {selectedInspection.attr?.inspector}
//                   </Typography>
//                   <Typography>
//                     <b>Date:</b>{" "}
//                     {selectedInspection.attr?.date
//                       ? new Date(selectedInspection.attr.date).toLocaleDateString()
//                       : "Unknown"}
//                   </Typography>
//                   <Typography sx={{ whiteSpace: "pre-wrap" }}>
//                     <b>Observations:</b>{" "}
//                     {selectedInspection.ai_notes?.inspectorObservations ||
//                       selectedInspection.notes?.inspectorObservations ||
//                       "No observations."}
//                   </Typography>
//                 </Stack>
//               </>
//             )}
//           </Paper>
//         </Modal>
//       </Box>
//     </Box>
//   );
// }



// // src/taskpane/components/main/Home.tsx
// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Divider,
//   List,
//   ListItemButton,
//   ListItemText,
//   Button,
//   Modal,
//   Paper,
//   Stack,
// } from "@mui/material";
// import { IconButton } from "@mui/material";
// import CloseIcon from "@mui/icons-material/Close";

// import AppHeader from "./AppHeader";
// import { fetchAddonReports, Inspection } from "../services/AddonReports";

// export default function Home() {
//   const [inspections, setInspections] = useState<Inspection[]>([]);
//   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string | null>(null);

//   const handleClose = () => setSelectedInspection(null);

//   const getReports = async () => {
//     setLoading(true);
//     setError(null);
//     try {
//       const data = await fetchAddonReports();
//       setInspections(data);
//     } catch (err: any) {
//       console.error(err);
//       setError(err.message);
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     getReports();
//   }, []);

//   return (
//     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
//       <AppHeader />

//       <Box display="flex" flexGrow={1}>
//         {/* Sidebar */}
//         <Box
//           p={2}
//           width={300}
//           bgcolor="#fff"
//           borderRight="1px solid #ddd"
//           display="flex"
//           flexDirection="column"
//         >
//           <Typography variant="h6" mb={1}>
//             Inspections
//           </Typography>
//           <Divider sx={{ mb: 2 }} />

//           {/* List */}
//           <List sx={{ flexGrow: 1, overflowY: "auto" }}>
//             {loading ? (
//               <Typography>Loading...</Typography>
//             ) : error ? (
//               <Typography color="error">{error}</Typography>
//             ) : inspections.length > 0 ? (
//               inspections.map((insp) => (
//                 <ListItemButton
//                   key={insp.id}
//                   onClick={() => setSelectedInspection(insp)}
//                   sx={{
//                     borderRadius: 1,
//                     mb: 1,
//                     "&:hover": { bgcolor: "#f0f0f0" },
//                   }}
//                 >
//                   <ListItemText
//                     primary={`#${insp.attr?.no || "N/A"} - ${insp.attr?.material || "Unknown Material"}`}
//                     secondary={`${insp.attr?.location || "Unknown"} • ${
//                       insp.attr?.inspector || "No Inspector"
//                     }`}
//                   />
//                 </ListItemButton>
//               ))
//             ) : (
//               <Typography>No inspections found.</Typography>
//             )}
//           </List>

//           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
//             Refresh Data
//           </Button>
//         </Box>

//         {/* Modal for Preview */}
//        {/* Modal for Preview */}
// {/* Modal for Preview */}
// <Modal
//   open={!!selectedInspection}
//   onClose={handleClose}
//   sx={{
//     display: "flex",
//     alignItems: "center",
//     justifyContent: "center",
//     p: 2,
//   }}
// >
//   <Paper
//     sx={{
//       position: "relative",
//       width: { xs: "100%", sm: "90%", md: 600 },
//       maxHeight: "86vh",
//       overflowY: "auto",
//       p: 3,
//       outline: "none",
//       borderRadius: 2,
//     }}
//   >
//     {/* Close Button (Top Right) */}
//     <IconButton
//       onClick={handleClose}
//       sx={{
//         position: "absolute",
//         top: 8,
//         right: 8,
//         color: "grey.600",
//       }}
//     >
//       <CloseIcon />
//     </IconButton>

//     {selectedInspection && (
//       <>
//         <Typography variant="h6" gutterBottom sx={{ pr: 5 }}>
//           Report #{selectedInspection.attr?.no}
//         </Typography>
//         <Divider sx={{ mb: 2 }} />

//         <Stack spacing={1}>
//           <Typography>
//             <b>Location:</b> {selectedInspection.attr?.location}
//           </Typography>
//           <Typography>
//             <b>Material:</b> {selectedInspection.attr?.material}
//           </Typography>
//           <Typography>
//             <b>Inspector:</b> {selectedInspection.attr?.inspector}
//           </Typography>
//           <Typography>
//             <b>Date:</b>{" "}
//             {selectedInspection.attr?.date
//               ? new Date(selectedInspection.attr.date).toLocaleDateString()
//               : "Unknown"}
//           </Typography>
//           <Typography sx={{ whiteSpace: "pre-wrap" }}>
//             <b>Observations:</b>{" "}
//             {selectedInspection.ai_notes?.inspectorObservations ||
//               selectedInspection.notes?.inspectorObservations ||
//               "No observations."}
//           </Typography>
//         </Stack>

//         <Button
//           onClick={handleClose}
//           variant="contained"
//           color="primary"
//           fullWidth
//           sx={{ mt: 3 }}
//         >
//           Close
//         </Button>
//       </>
//     )}
//   </Paper>
// </Modal>


//       </Box>
//     </Box>
//   );
// }


// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Divider,
//   Button,
//   Modal,
//   Paper,
//   Stack,
//   IconButton,
//   Card,
//   CardContent,
//   CardActions,
//   CircularProgress,
//   Snackbar,
//   Alert,
// } from "@mui/material";
// import CloseIcon from "@mui/icons-material/Close";
// import AppHeader from "./AppHeader";
// import { fetchAddonReports, Inspection } from "../services/AddonReports";
// import stringSimilarity from "string-similarity";

// export default function Home() {
//   const [inspections, setInspections] = useState<Inspection[]>([]);
//   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string | null>(null);
//   const [message, setMessage] = useState<string | null>(null);

//   // ---------------------------
//   // Fetch Addon Reports
//   // ---------------------------
//   const getReports = async () => {
//     setLoading(true);
//     setError(null);
//     try {
//       const data = await fetchAddonReports();
//       setInspections(data);
//       setMessage("Inspection reports loaded successfully!");
//     } catch (err: any) {
//       console.error(err);
//       setError("Failed to fetch inspection reports. Please try again.");
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     getReports();
//   }, []);

//   // ---------------------------
//   // Client-specific section titles
//   // ---------------------------
//   const templatesByClient: Record<string, string[]> = {
//     WoodTech: [
//       "Introduction / Scope",
//       "Purpose of Inspection",
//       "Reference Documents and Standards",
//       "Findings",
//       "Photos",
//     ],
//     SteelWorks: ["Inspection Overview", "Observations", "Defects", "Conclusion"],
//   };

//   // ---------------------------
//   // Insert into Word
//   // ---------------------------
//   // const insertIntoWord = async (inspection: Inspection) => {
//   //   const clientName = inspection.attr?.client || "WoodTech";
//   //   const clientSections = templatesByClient[clientName] || templatesByClient["WoodTech"];

//   //   const reportSections = ["Purpose of Inspection", "Findings", "Recommendations", "Photos"];

//   //   const matches = reportSections.map((section) => {
//   //     const { bestMatch } = stringSimilarity.findBestMatch(section, clientSections);
//   //     return {
//   //       reportSection: section,
//   //       matchedTo: bestMatch.target,
//   //     };
//   //   });

//   //   try {
//   //     await Word.run(async (context) => {
//   //       const paragraphs = context.document.body.paragraphs;
//   //       paragraphs.load("items");
//   //       await context.sync();

//   //       for (const match of matches) {
//   //         const paragraph = paragraphs.items.find((p) =>
//   //           p.text.toLowerCase().includes(match.matchedTo.toLowerCase())
//   //         );

//   //         if (paragraph) {
//   //           paragraph.insertParagraph(
//   //             `${match.reportSection}: ${
//   //               inspection.ai_notes?.inspectorObservations || "No data available"
//   //             }`,
//   //             "After"
//   //           );
//   //         }
//   //       }

//   //       await context.sync();
//   //     });

//   //     setMessage(`✅ Report inserted successfully for ${cliyentName}!`);
//   //   } catch (error) {
//   //     console.error("Error inserting into Word:", error);
//   //     setError("⚠️ Failed to insert report into Word.");
//   //   }
//   // };



// const getselectedInspection = (inspection:Inspection) => {
// console.log("Selected Inspection:", inspection);
// }

// // place holders

// const insertIntoWord = async (inspection: any) => {

//   console.log("Inserting inspection into Word:", inspection);
//   try {
//     // 🧹 Clean and normalize API text
//     const cleanApiText = (text: unknown): string => {
//       if (typeof text !== "string") return "";
//       let textStr = text;

//       try {
//         // Decode double-escaped newlines (e.g. \\n\\n)
//         if (textStr.includes("\\n")) {
//           textStr = JSON.parse(`"${textStr.replace(/"/g, '\\"')}"`);
//         }
//       } catch {
//         // Fail silently
//       }

//       let cleaned = textStr
//         .replace(/(\\r|\\n|\\\\n)+/g, "\n") // Normalize newlines
//         .replace(/\*\*(.*?)\*\*/g, "$1") // Remove markdown bold
//         .replace(/\n\s*-\s*/g, "\n• ") // Convert dashes to bullets
//         .replace(/•+/g, "•") // Fix multiple bullets
//         .replace(/\.\s*\n/g, ". ") // Fix dot-newline spacing
//         .replace(/\n{3,}/g, "\n\n") // Collapse multiple blank lines
//         .trim();

//       return cleaned;
//     };

//     // 🔍 Recursively extract all text fields from the API response
//     const fields: Record<string, string> = {};
//     const extractFields = (obj: any, prefix = "") => {
//       console.log("Extracting fields from object:", prefix);
//       if (!obj || typeof obj !== "object") return;
//       for (const [key, value] of Object.entries(obj)) {
//         if (typeof value === "object") extractFields(value, key);
//         else if (typeof value === "string" && value.trim() !== "") {
//           fields[key] = cleanApiText(value);
//         }
//       }
//     };

//     extractFields(inspection);
//     extractFields(inspection.attr);
//     extractFields(inspection.notes);
//     extractFields(inspection.ai_notes);
//     extractFields(inspection.custom_attr);

//     console.log("🧩 Detected placeholders:", Object.keys(fields));

//     // 🧠 Replace placeholders in Word
//     await Word.run(async (context) => {
//       const body = context.document.body;

//       for (const [key, value] of Object.entries(fields)) {
//         const placeholder = `{{${key}}}`;

//         const searchResults = body.search(placeholder, {
//           matchCase: false,
//           matchWholeWord: false,
//         });

//         searchResults.load("items");
//         await context.sync();

//         if (searchResults.items.length > 0) {
//           console.log(`✏️ Replacing placeholder: ${placeholder}`);
//           for (const range of searchResults.items) {
//             range.insertText(value, "Replace");
//           }
//         }
//       }

//       await context.sync();
//     });

//     console.log("✅ Dynamic report inserted successfully!");
//   } catch (error) {
//     console.error("⚠️ Error inserting into Word:", error);
//   }
// };



// // const insertIntoWord = async (inspection: Inspection) => {
// //   try {
// // const cleanApiText = (text: unknown): string => {
// //   if (typeof text !== "string") return "";
// //   let textStr = text;

// //   try {
// //     // Decode double-escaped newlines (e.g. \\n\\n)
// //     if (textStr.includes("\\n")) {
// //       textStr = JSON.parse(`"${textStr.replace(/"/g, '\\"')}"`);
// //     }
// //   } catch {
// //     // Fail silently
// //   }

// //   let cleaned = textStr
// //     // Normalize CRLFs and escaped newlines
// //     .replace(/(\\r|\\n|\\\\n)+/g, "\n")

// //     // Remove markdown bold markers
// //     .replace(/\*\*(.*?)\*\*/g, "$1")

// //     // Replace hyphens used for bullets
// //     .replace(/\n\s*-\s*/g, "\n• ")

// //     // Replace multiple bullet dots (••) with one
// //     .replace(/•+/g, "•")

// //     // Remove stray periods before newlines (".\n" → ". ")
// //     .replace(/\.\s*\n/g, ". ")

// //     // Collapse multiple newlines (more than 2 → 2)
// //     .replace(/\n{3,}/g, "\n\n")

// //     // Trim whitespace
// //     .trim();

// //   return cleaned;
// // };



// //     // ✅ Helper: Flatten all nested fields from the API object
// //     const fields: Record<string, string> = {};
// //     const extractFields = (obj: any, prefix = "") => {
// //       if (!obj || typeof obj !== "object") return;
// //       for (const [key, value] of Object.entries(obj)) {
// //         const fullKey = prefix ? `${prefix}.${key}` : key;
// //         if (typeof value === "object") extractFields(value, fullKey);
// //         else if (typeof value === "string" && value.trim() !== "") {
// //           fields[key] = cleanApiText(value);
// //         }
// //       }
// //     };

// //     extractFields(inspection);
// //     extractFields(inspection.attr);
// //     extractFields(inspection.notes);
// //     extractFields(inspection.ai_notes);
// //     extractFields(inspection.custom_attr);

// //     const fieldKeys = Object.keys(fields);
// //     console.log("Detected report fields:", fieldKeys);

// //     // ✅ Insert into Word
// //     await Word.run(async (context) => {
// //       const paragraphs = context.document.body.paragraphs;
// //       paragraphs.load("items");
// //       await context.sync();

// //       const paraItems = paragraphs.items.map((p) => p.text.toLowerCase());

// //       for (const key of fieldKeys) {
// //         // find best matching paragraph in Word
// //         const { bestMatch } = stringSimilarity.findBestMatch(
// //           key.toLowerCase(),
// //           paraItems
// //         );
// //         const bestIndex = paraItems.findIndex(
// //           (p) => p === bestMatch.target
// //         );

// //         if (bestIndex >= 0 && bestMatch.rating > 0.5) {
// //           const paragraph = paragraphs.items[bestIndex];
// //           const value = fields[key];

// //           // ✅ Insert heading + cleaned value
// //           if (value.includes("\n")) {
// //             const [firstLine, ...rest] = value.split("\n\n");
// //             const headingParagraph = paragraph.insertParagraph(firstLine, "After");
// //             headingParagraph.font.bold = true;

// //             paragraph.insertParagraph(rest.join("\n\n"), "After");
// //           } else {
// //             paragraph.insertParagraph(value, "After");
// //           }
// //         }
// //       }

// //       await context.sync();
// //     });

// //     setMessage("✅ Dynamic report inserted successfully!");
// //   } catch (error) {
// //     console.error("Error inserting into Word:", error);
// //     setError("⚠️ Failed to insert dynamic report into Word.");
// //   }
// // };


// // async function insertIntoWord(inspection: any) {
// //   await Word.run(async (context) => {
// //     const body = context.document.body;

// //     // Extract key values safely
// //     const obs =
// //       inspection.inspectorObservations?.trim() ||
// //       inspection.transcript?.trim() ||
// //       "No observations available";

// //     const refDocs =
// //       inspection.referenceDocuments?.trim() || "No reference documents available";

// //     const mdb =
// //       inspection.custom_attr?.["MDB Completion"] ||
// //       "No MDB Completion status provided";

// //     // Map of placeholders to values
// //     const replacements: Record<string, string> = {
// //       inspectorObservations: obs,
// //       referenceDocuments: refDocs,
// //       "MDB Completion": mdb,
// //       Summary:
// //         "Overall, the inspection has been completed. Please review observations for non-conformances.",
// //     };

// //     // Loop and replace placeholders
// //     for (const [key, value] of Object.entries(replacements)) {
// //       const searchResults = body.search("(Content from API will be inserted here)", {
// //         matchCase: false,
// //         matchWholeWord: false,
// //       });
// //       searchResults.load("items");
// //       await context.sync();

// //       for (const range of searchResults.items) {
// //         const prev = range.paragraphs.getFirst().getPrevious();
// //         prev.load("text");
// //         await context.sync();

// //         if (prev.text.includes(key)) {
// //           range.insertText(String(value), "Replace");
// //         }
// //       }
// //     }

// //     // // Add signature (optional)
// //     // if (inspection.signature?.publicUrl) {
// //     //   // const base64 = await fetchAsBase64(inspection.signature.publicUrl);
// //     //   body.insertParagraph("Inspector Signature:", "End");
// //     //   // body.insertInlinePictureFromBase64(base64, "End");
// //     // }

// //     body.insertParagraph("\n--- End of Report ---\n", "End");

// //     await context.sync();
// //   });
// // }



//   // ---------------------------
//   // Render UI
//   // ---------------------------




//   return (
//     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
//       <AppHeader />

//       {/* Loading Overlay */}
//       {loading && (
//         <Box
//           position="fixed"
//           top={0}
//           left={0}
//           right={0}
//           bottom={0}
//           display="flex"
//           alignItems="center"
//           justifyContent="center"
//           bgcolor="rgba(255,255,255,0.7)"
//           zIndex={9999}
//         >
//           <CircularProgress size={60} />
//         </Box>
//       )}

//       {/* Snackbar Alerts */}
//       <Snackbar
//         open={!!message || !!error}
//         autoHideDuration={4000}
//         onClose={() => {
//           setMessage(null);
//           setError(null);
//         }}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         {message ? (
//           <Alert severity="success" sx={{ width: "100%" }}>
//             {message}
//           </Alert>
//         ) : error ? (
//           <Alert severity="error" sx={{ width: "100%" }}>
//             {error}
//           </Alert>
//         ) : null}
//       </Snackbar>

//       <Box display="flex" flexGrow={1}>
//         {/* Sidebar */}
//         <Box
//           p={3}
//           width={340}
//           bgcolor="#fff"
//           borderRight="1px solid #ddd"
//           display="flex"
//           flexDirection="column"
//           sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
//         >

//           <Divider sx={{ mb: 2 }} />

//           {/* List of Inspections */}
//           <Box flexGrow={1} overflow="auto" pr={1}>
//             {inspections.length > 0 ? (
//               inspections.map((insp) => (
//                 <Card
//                   key={insp.id}
//                   sx={{
//                     mb: 2,
//                     borderRadius: 2,
//                     boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
//                     transition: "0.3s",
//                     "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
//                   }}
//                 >
//                   <CardContent>
//                     <Typography variant="subtitle1" fontWeight={600}>
//                       #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
//                     </Typography>
//                     <Typography variant="body2" color="text.secondary">
//                       {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
//                     </Typography>
//                   </CardContent>

//                   <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
//                     <Button
//                       variant="outlined"
//                       size="small"
//                       onClick={() => getselectedInspection(insp)}
//                     >
//                       Preview
//                     </Button>
//                     <Button
//                       variant="contained"
//                       size="small"
//                       onClick={() => insertIntoWord(insp)}
//                     >
//                       Insert
//                     </Button>
//                   </CardActions>
//                 </Card>
//               ))
//             ) : (
//               !loading && <Typography>No inspections found.</Typography>
//             )}
//           </Box>

//           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
//             Refresh Data
//           </Button>
//         </Box>

//         {/* Modal for Preview */}
//         <Modal
//           open={!!selectedInspection}
//           onClose={() => setSelectedInspection(null)}
//           sx={{
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "center",
//             p: 2,
//           }}
//         >
//           <Paper
//             sx={{
//               position: "relative",
//               width: { xs: "100%", sm: "90%", md: 600 },
//               maxHeight: "86vh",
//               overflowY: "auto",
//               p: 3,
//               borderRadius: 3,
//               boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
//             }}
//           >
//             <IconButton
//               onClick={() => setSelectedInspection(null)}
//               sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
//             >
//               <CloseIcon />
//             </IconButton>

//             {selectedInspection && (
//               <>
//                 <Typography variant="h6" gutterBottom>
//                   Report #{selectedInspection.attr?.no}
//                 </Typography>
//                 <Divider sx={{ mb: 2 }} />
//                 <Stack spacing={1}>
//                   <Typography>
//                     <b>Location:</b> {selectedInspection.attr?.location}
//                   </Typography>
//                   <Typography>
//                     <b>Material:</b> {selectedInspection.attr?.material}
//                   </Typography>
//                   <Typography>
//                     <b>Inspector:</b> {selectedInspection.attr?.inspector}
//                   </Typography>
//                   <Typography>
//                     <b>Date:</b>{" "}
//                     {selectedInspection.attr?.date
//                       ? new Date(selectedInspection.attr.date).toLocaleDateString()
//                       : "Unknown"}
//                   </Typography>
//                   <Typography sx={{ whiteSpace: "pre-wrap" }}>
//                     <b>Observations:</b>{" "}
//                     {selectedInspection.ai_notes?.inspectorObservations ||
//                       selectedInspection.notes?.inspectorObservations ||
//                       "No observations."}
//                   </Typography>
//                 </Stack>
//               </>
//             )}
//           </Paper>
//         </Modal>
//       </Box>
//     </Box>
//   );
// }

// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Divider,
//   Button,
//   Modal,
//   Paper,
//   Stack,
//   IconButton,
//   Card,
//   CardContent,
//   CardActions,
//   CircularProgress,
//   Snackbar,
//   Alert,
// } from "@mui/material";
// import CloseIcon from "@mui/icons-material/Close";
// import AddIcon from "@mui/icons-material/Add";
// import AppHeader from "./AppHeader";
// import { fetchAddonReports, Inspection } from "../services/AddonReports";

// export default function Home() {
//   const [inspections, setInspections] = useState<Inspection[]>([]);
//   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string | null>(null);
//   const [message, setMessage] = useState<string | null>(null);
//   const [expandedFields, setExpandedFields] = useState<Record<string, boolean>>({});

//   // ---------------------------
//   // Fetch Addon Reports
//   // ---------------------------
//   const getReports = async () => {
//     setLoading(true);
//     setError(null);
//     try {
//       const data = await fetchAddonReports();
//       setInspections(data);
//       setMessage("Inspection reports loaded successfully!");
//     } catch (err: any) {
//       console.error(err);
//       setError("Failed to fetch inspection reports. Please try again.");
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     getReports();
//   }, []);

//   // ---------------------------
//   // Helper: Flatten nested objects
//   // ---------------------------
//   const flattenInspection = (obj: any, prefix = ""): Record<string, any> => {
//     let result: Record<string, any> = {};
//     for (const [key, value] of Object.entries(obj || {})) {
//       const fullKey = prefix ? `${prefix}.${key}` : key;
//       if (typeof value === "object" && value !== null) {
//         result = { ...result, ...flattenInspection(value, fullKey) };
//       } else {
//         result[fullKey] = value;
//       }
//     }
//     return result;
//   };

//   // ---------------------------
//   // Insert one field at selection in Word
//   // ---------------------------
//   const insertFieldAtSelection = async (key: string, value: string) => {
//     try {
//       await Word.run(async (context) => {
//         const range = context.document.getSelection();

//         // ✅ Check if value is an image URL
//         if (
//           /^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(value) ||
//           value.includes("supabase.co/storage/")
//         ) {
//           const response = await fetch(value);
//           const blob = await response.blob();
//           const base64 = await blobToBase64(blob);

//           range.insertInlinePictureFromBase64(base64, "Replace");
//           await context.sync();

//           setMessage(`🖼️ Inserted image from ${key}`);
//           return;
//         }

//         // ✅ Otherwise treat as text
//         const paragraphs = value
//           .replace(/\r/g, "")
//           .split(/\n{2,}/)
//           .filter((p) => p.trim() !== "");

//         paragraphs.forEach((para) => {
//           const trimmed = para.trim();

//           if (/^[-•]\s/.test(trimmed)) {
//             const lines = trimmed
//               .split(/\n|(?=-\s)/)
//               .map((line) => line.replace(/^[-•]\s*/, "").trim())
//               .filter((line) => line.length > 0);

//             const bulletRange = range.insertParagraph("", "After");
//             lines.forEach((line) => {
//               const p = bulletRange.insertParagraph("", "After");
//               // Insert a bullet character at the start and then the formatted text (avoid using non-existent enum)
//               p.insertText("• ", "Start");
//               insertFormattedText(p, line);
//             });
//           } else {
//             const p = range.insertParagraph("", "After");
//             insertFormattedText(p, trimmed);
//           }
//         });

//         await context.sync();
//       });

//       setMessage(`✅ Inserted "${key}" value into document`);
//     } catch (err) {
//       console.error(err);
//       setError("⚠️ Failed to insert field into Word document");
//     }
//   };

//   // ---------------------------
//   // Helper: Blob → Base64
//   // ---------------------------
//   function blobToBase64(blob: Blob): Promise<string> {
//     return new Promise((resolve, reject) => {
//       const reader = new FileReader();
//       reader.onloadend = () => {
//         const base64data = (reader.result as string).split(",")[1];
//         resolve(base64data);
//       };
//       reader.onerror = reject;
//       reader.readAsDataURL(blob);
//     });
//   }

//   // ---------------------------
//   // Helper: Bold Markdown (**text**)
//   // ---------------------------
//   function insertFormattedText(paragraph: Word.Paragraph, text: string) {
//     const boldRegex = /\*\*(.*?)\*\*/g;
//     let match;
//     let lastIndex = 0;

//     while ((match = boldRegex.exec(text)) !== null) {
//       const beforeText = text.substring(lastIndex, match.index);
//       if (beforeText) paragraph.insertText(beforeText, "End");

//       const boldText = match[1];
//       const boldRange = paragraph.insertText(boldText, "End");
//       boldRange.font.bold = true;

//       lastIndex = match.index + match[0].length;
//     }

//     const remaining = text.substring(lastIndex);
//     if (remaining) paragraph.insertText(remaining, "End");
//   }

//   // ---------------------------
//   // UI Rendering
//   // ---------------------------
//   return (
//     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
//       <AppHeader />

//       {/* Loading */}
//       {loading && (
//         <Box
//           position="fixed"
//           top={0}
//           left={0}
//           right={0}
//           bottom={0}
//           display="flex"
//           alignItems="center"
//           justifyContent="center"
//           bgcolor="rgba(255,255,255,0.7)"
//           zIndex={9999}
//         >
//           <CircularProgress size={60} />
//         </Box>
//       )}

//       {/* Snackbar Alerts */}
//       <Snackbar
//         open={!!message || !!error}
//         autoHideDuration={4000}
//         onClose={() => {
//           setMessage(null);
//           setError(null);
//         }}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         {message ? (
//           <Alert severity="success" sx={{ width: "100%" }}>
//             {message}
//           </Alert>
//         ) : error ? (
//           <Alert severity="error" sx={{ width: "100%" }}>
//             {error}
//           </Alert>
//         ) : null}
//       </Snackbar>

//       <Box display="flex" flexGrow={1}>
//         {/* Sidebar */}
//         <Box
//           p={3}
//           width={340}
//           bgcolor="#fff"
//           borderRight="1px solid #ddd"
//           display="flex"
//           flexDirection="column"
//           sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
//         >
//           <Divider sx={{ mb: 2 }} />

//           {/* List of Inspections */}
//           <Box flexGrow={1} overflow="auto" pr={1}>
//             {inspections.length > 0 ? (
//               inspections.map((insp) => (
//                 <Card
//                   key={insp.id}
//                   sx={{
//                     mb: 2,
//                     borderRadius: 2,
//                     boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
//                     transition: "0.3s",
//                     "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
//                   }}
//                 >
//                   <CardContent>
//                     <Typography variant="subtitle1" fontWeight={600}>
//                       #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
//                     </Typography>
//                     <Typography variant="body2" color="text.secondary">
//                       {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
//                     </Typography>
//                   </CardContent>

//                   <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
//                     <Button
//                       variant="outlined"
//                       size="small"
//                       onClick={() => setSelectedInspection(insp)}
//                     >
//                       Preview
//                     </Button>
//                   </CardActions>
//                 </Card>
//               ))
//             ) : (
//               !loading && <Typography>No inspections found.</Typography>
//             )}
//           </Box>

//           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
//             Refresh Data
//           </Button>
//         </Box>

//         {/* Modal for Preview */}
//         <Modal
//           open={!!selectedInspection}
//           onClose={() => setSelectedInspection(null)}
//           sx={{
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "center",
//             p: 2,
//           }}
//         >
//           <Paper
//             sx={{
//               position: "relative",
//               width: { xs: "100%", sm: "90%", md: 700 },
//               maxHeight: "86vh",
//               overflowY: "auto",
//               p: 3,
//               borderRadius: 3,
//               boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
//             }}
//           >
//             <IconButton
//               onClick={() => setSelectedInspection(null)}
//               sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
//             >
//               <CloseIcon />
//             </IconButton>

//             {selectedInspection && (
//               <>
//                 <Typography variant="h6" gutterBottom>
//                   Inspection Report #{selectedInspection.attr?.no || "N/A"}
//                 </Typography>
//                 <Divider sx={{ mb: 2 }} />

//                 <Stack spacing={1}>
//                   {Object.entries(flattenInspection(selectedInspection))
//                     .filter(([key]) => !["id", "cid", "uid", "created_at", "attr.no"].includes(key))
//                     .map(([key, value]) => {
//                       const valStr = String(value || "");
//                       const isLong = valStr.length > 150;
//                       const isExpanded = expandedFields[key];
//                       const displayValue = isExpanded
//                         ? valStr
//                         : isLong
//                         ? `${valStr.substring(0, 150)}...`
//                         : valStr;

//                       return (
//                         <Box
//                           key={key}
//                           display="flex"
//                           alignItems="flex-start"
//                           justifyContent="space-between"
//                           bgcolor="#f9f9f9"
//                           borderRadius={1.5}
//                           px={2}
//                           py={1.5}
//                         >
//                           <Box flex={1} mr={1}>
//                             <Typography variant="subtitle2" fontWeight={600}>
//                               {key}
//                             </Typography>

//                             {/* Show image preview if URL is image */}
//                             {/^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(valStr) ||
//                             valStr.includes("supabase.co/storage/") ? (
//                               <Box mt={1}>
//                                 <img
//                                   src={valStr}
//                                   alt={key}
//                                   style={{
//                                     maxWidth: "100%",
//                                     maxHeight: "150px",
//                                     borderRadius: "8px",
//                                   }}
//                                 />
//                               </Box>
//                             ) : (
//                               <Typography
//                                 variant="body2"
//                                 color="text.secondary"
//                                 sx={{ whiteSpace: "pre-wrap" }}
//                               >
//                                 {displayValue}
//                               </Typography>
//                             )}

//                             {isLong && (
//                               <Button
//                                 size="small"
//                                 variant="text"
//                                 onClick={() =>
//                                   setExpandedFields((prev) => ({
//                                     ...prev,
//                                     [key]: !prev[key],
//                                   }))
//                                 }
//                                 sx={{ mt: 0.5 }}
//                               >
//                                 {isExpanded ? "Show less" : "Show more"}
//                               </Button>
//                             )}
//                           </Box>

//                           {/* ✅ Insert button beside each field */}
//                           <IconButton
//                             color="primary"
//                             onClick={() => insertFieldAtSelection(key, valStr)}
//                             title="Insert this field into Word"
//                           >
//                             <AddIcon />
//                           </IconButton>
//                         </Box>
//                       );
//                     })}
//                 </Stack>
//               </>
//             )}
//           </Paper>
//         </Modal>
//       </Box>
//     </Box>
//   );
// }


//// above code is working but inserting next to selection








// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Divider,
//   Button,
//   Modal,
//   Paper,
//   Stack,
//   IconButton,
//   Card,
//   CardContent,
//   CardActions,
//   CircularProgress,
//   Snackbar,
//   Alert,
// } from "@mui/material";
// import CloseIcon from "@mui/icons-material/Close";
// import AddIcon from "@mui/icons-material/Add";
// import AppHeader from "./AppHeader";
// import { fetchAddonReports, Inspection } from "../services/AddonReports";

// export default function Home() {
//   const [inspections, setInspections] = useState<Inspection[]>([]);
//   const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string | null>(null);
//   const [message, setMessage] = useState<string | null>(null);
//   const [expandedFields, setExpandedFields] = useState<Record<string, boolean>>({});

//   // ---------------------------
//   // Fetch Addon Reports
//   // ---------------------------
//   const getReports = async () => {
//     setLoading(true);
//     setError(null);
//     try {
//       const data = await fetchAddonReports();
//       setInspections(data);
//       setMessage("Inspection reports loaded successfully!");
//     } catch (err: any) {
//       console.error(err);
//       setError("Failed to fetch inspection reports. Please try again.");
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     getReports();
//   }, []);

//   // ---------------------------
//   // Helper: Flatten nested objects
//   // ---------------------------
//   const flattenInspection = (obj: any, prefix = ""): Record<string, any> => {
//     let result: Record<string, any> = {};
//     for (const [key, value] of Object.entries(obj || {})) {
//       const fullKey = prefix ? `${prefix}.${key}` : key;
//       if (typeof value === "object" && value !== null) {
//         result = { ...result, ...flattenInspection(value, fullKey) };
//       } else {
//         result[fullKey] = value;
//       }
//     }
//     return result;
//   };

//  const insertFieldAtSelection = async (key: string, value: string) => {
//   try {
//     await Word.run(async (context) => {
//       let range = context.document.getSelection();

//       // ✅ Handle image URLs
//       if (
//         /^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(value) ||
//         value.includes("supabase.co/storage/")
//       ) {
//         const response = await fetch(value);
//         const blob = await response.blob();
//         const base64 = await blobToBase64(blob);

//         range.insertInlinePictureFromBase64(base64, "Replace");
//         await context.sync();

//         setMessage(`🖼️ Inserted image from ${key}`);
//         return;
//       }

//       // ✅ Handle text content
//       const paragraphs = value
//         .replace(/\r/g, "")
//         .split(/\n{2,}/)
//         .filter((p) => p.trim() !== "");

//       // Clear selection before inserting
//       range.insertText("", "Replace");

//       for (let i = 0; i < paragraphs.length; i++) {
//         const para = paragraphs[i].trim();

//         if (/^[-•]\s/.test(para)) {
//           // bullet list
//           const lines = para
//             .split(/\n|(?=-\s)/)
//             .map((line) => line.replace(/^[-•]\s*/, "").trim())
//             .filter((line) => line.length > 0);

//         for (const line of lines) {
//   const p = range.insertParagraph("", "After");
//   insertFormattedText(p, line);
// p.style = "List Paragraph";

//   range = p.getRange("End");
// }

//         } else {
//           const p = range.insertParagraph("", "After");
//           insertFormattedText(p, para);
//           range = p.getRange("End"); // ✅ convert Paragraph → Range
//         }

//         if (i < paragraphs.length - 1) {
//           range.insertParagraph("", "After");
//         }
//       }

//       await context.sync();
//     });

//     setMessage(`✅ Inserted "${key}" value into document`);
//   } catch (err) {
//     console.error(err);
//     setError("⚠️ Failed to insert field into Word document");
//   }
// };

// // Helper: convert Blob → Base64
// function blobToBase64(blob: Blob): Promise<string> {
//   return new Promise((resolve, reject) => {
//     const reader = new FileReader();
//     reader.onloadend = () => {
//       const base64data = (reader.result as string).split(",")[1];
//       resolve(base64data);
//     };
//     reader.onerror = reject;
//     reader.readAsDataURL(blob);
//   });
// }


//   // ---------------------------
//   // Helper: Bold Markdown (**text**)
//   // ---------------------------
//   function insertFormattedText(paragraph: Word.Paragraph, text: string) {
//     const boldRegex = /\*\*(.*?)\*\*/g;
//     let match;
//     let lastIndex = 0;

//     while ((match = boldRegex.exec(text)) !== null) {
//       const beforeText = text.substring(lastIndex, match.index);
//       if (beforeText) paragraph.insertText(beforeText, "End");

//       const boldText = match[1];
//       const boldRange = paragraph.insertText(boldText, "End");
//       boldRange.font.bold = true;

//       lastIndex = match.index + match[0].length;
//     }

//     const remaining = text.substring(lastIndex);
//     if (remaining) paragraph.insertText(remaining, "End");
//   }

//   // ---------------------------
//   // UI Rendering
//   // ---------------------------
//   return (
//     <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
//       <AppHeader />

//       {/* Loading */}
//       {loading && (
//         <Box
//           position="fixed"
//           top={0}
//           left={0}
//           right={0}
//           bottom={0}
//           display="flex"
//           alignItems="center"
//           justifyContent="center"
//           bgcolor="rgba(255,255,255,0.7)"
//           zIndex={9999}
//         >
//           <CircularProgress size={60} />
//         </Box>
//       )}

//       {/* Snackbar Alerts */}
//       <Snackbar
//         open={!!message || !!error}
//         autoHideDuration={4000}
//         onClose={() => {
//           setMessage(null);
//           setError(null);
//         }}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         {message ? (
//           <Alert severity="success" sx={{ width: "100%" }}>
//             {message}
//           </Alert>
//         ) : error ? (
//           <Alert severity="error" sx={{ width: "100%" }}>
//             {error}
//           </Alert>
//         ) : null}
//       </Snackbar>

//       <Box display="flex" flexGrow={1}>
//         {/* Sidebar */}
//         <Box
//           p={3}
//           width={340}
//           bgcolor="#fff"
//           borderRight="1px solid #ddd"
//           display="flex"
//           flexDirection="column"
//           sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
//         >
//           <Divider sx={{ mb: 2 }} />

//           {/* List of Inspections */}
//           <Box flexGrow={1} overflow="auto" pr={1}>
//             {inspections.length > 0 ? (
//               inspections.map((insp) => (
//                 <Card
//                   key={insp.id}
//                   sx={{
//                     mb: 2,
//                     borderRadius: 2,
//                     boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
//                     transition: "0.3s",
//                     "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
//                   }}
//                 >
//                   <CardContent>
//                     <Typography variant="subtitle1" fontWeight={600}>
//                       #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
//                     </Typography>
//                     <Typography variant="body2" color="text.secondary">
//                       {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
//                     </Typography>
//                   </CardContent>

//                   <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
//                     <Button
//                       variant="outlined"
//                       size="small"
//                       onClick={() => setSelectedInspection(insp)}
//                     >
//                       Preview
//                     </Button>
//                   </CardActions>
//                 </Card>
//               ))
//             ) : (
//               !loading && <Typography>No inspections found.</Typography>
//             )}
//           </Box>

//           <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
//             Refresh Data
//           </Button>
//         </Box>

//         {/* Modal for Preview */}
//         <Modal
//           open={!!selectedInspection}
//           onClose={() => setSelectedInspection(null)}
//           sx={{
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "center",
//             p: 2,
//           }}
//         >
//           <Paper
//             sx={{
//               position: "relative",
//               width: { xs: "100%", sm: "90%", md: 700 },
//               maxHeight: "86vh",
//               overflowY: "auto",
//               p: 3,
//               borderRadius: 3,
//               boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
//             }}
//           >
//             <IconButton
//               onClick={() => setSelectedInspection(null)}
//               sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
//             >
//               <CloseIcon />
//             </IconButton>

//             {selectedInspection && (
//               <>
//                 <Typography variant="h6" gutterBottom>
//                   Inspection Report #{selectedInspection.attr?.no || "N/A"}
//                 </Typography>
//                 <Divider sx={{ mb: 2 }} />

//                 <Stack spacing={1}>
//                   {Object.entries(flattenInspection(selectedInspection))
//                     .filter(([key]) => !["id", "cid", "uid", "created_at", "attr.no"].includes(key))
//                     .map(([key, value]) => {
//                       const valStr = String(value || "");
//                       const isLong = valStr.length > 150;
//                       const isExpanded = expandedFields[key];
//                       const displayValue = isExpanded
//                         ? valStr
//                         : isLong
//                         ? `${valStr.substring(0, 150)}...`
//                         : valStr;

//                       return (
//                         <Box
//                           key={key}
//                           display="flex"
//                           alignItems="flex-start"
//                           justifyContent="space-between"
//                           bgcolor="#f9f9f9"
//                           borderRadius={1.5}
//                           px={2}
//                           py={1.5}
//                         >
//                           <Box flex={1} mr={1}>
//                             <Typography variant="subtitle2" fontWeight={600}>
//                               {key}
//                             </Typography>

//                             {/* Show image preview if URL is image */}
//                             {/^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(valStr) ||
//                             valStr.includes("supabase.co/storage/") ? (
//                               <Box mt={1}>
//                                 <img
//                                   src={valStr}
//                                   alt={key}
//                                   style={{
//                                     maxWidth: "100%",
//                                     maxHeight: "150px",
//                                     borderRadius: "8px",
//                                   }}
//                                 />
//                               </Box>
//                             ) : (
//                               <Typography
//                                 variant="body2"
//                                 color="text.secondary"
//                                 sx={{ whiteSpace: "pre-wrap" }}
//                               >
//                                 {displayValue}
//                               </Typography>
//                             )}

//                             {isLong && (
//                               <Button
//                                 size="small"
//                                 variant="text"
//                                 onClick={() =>
//                                   setExpandedFields((prev) => ({
//                                     ...prev,
//                                     [key]: !prev[key],
//                                   }))
//                                 }
//                                 sx={{ mt: 0.5 }}
//                               >
//                                 {isExpanded ? "Show less" : "Show more"}
//                               </Button>
//                             )}
//                           </Box>

//                           {/* ✅ Insert button beside each field */}
//                           <IconButton
//                             color="primary"
//                             onClick={() => insertFieldAtSelection(key, valStr)}
//                             title="Insert this field into Word"
//                           >
//                             <AddIcon />
//                           </IconButton>
//                         </Box>
//                       );
//                     })}
//                 </Stack>
//               </>
//             )}
//           </Paper>
//         </Modal>
//       </Box>
//     </Box>
//   );
// }

/// the abve code is working and inserting at selection but it is need ui improvement

import React, { useEffect, useState } from "react";
import {
  Box,
  Typography,
  Divider,
  Button,
  Modal,
  Paper,
  Stack,
  IconButton,
  Card,
  CardContent,
  CardActions,
  CircularProgress,
  Snackbar,
  Alert,
} from "@mui/material";
import CloseIcon from "@mui/icons-material/Close";
import AddIcon from "@mui/icons-material/Add";
import AppHeader from "./AppHeader";
import { fetchAddonReports, Inspection } from "../services/AddonReports";

export default function Home() {
  const [inspections, setInspections] = useState<Inspection[]>([]);
  const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [message, setMessage] = useState<string | null>(null);
  const [expandedFields, setExpandedFields] = useState<Record<string, boolean>>({});

  // ---------------------------
  // Fetch Addon Reports
  // ---------------------------
  const getReports = async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await fetchAddonReports();
      setInspections(data);
      setMessage("Inspection reports loaded successfully!");
    } catch (err: any) {
      console.error(err);
      setError("Failed to fetch inspection reports. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    getReports();
  }, []);

  // ---------------------------
  // Helper: Flatten nested objects
  // ---------------------------
  const flattenInspection = (obj: any, prefix = ""): Record<string, any> => {
    let result: Record<string, any> = {};
    for (const [key, value] of Object.entries(obj || {})) {
      const fullKey = prefix ? `${prefix}.${key}` : key;
      if (typeof value === "object" && value !== null) {
        result = { ...result, ...flattenInspection(value, fullKey) };
      } else {
        result[fullKey] = value;
      }
    }
    return result;
  };

  // ---------------------------
  // Helper: Make field names readable
  // ---------------------------



  // ---------------------------
  // Helper: Make field names readable
  // ---------------------------
  function formatKeyLabel(key: string): string {
    // Remove known prefixes like attr. or custom_attr.
    key = key.replace(/^(attr\.|custom_attr\.)/, "");

    const parts = key.split(".");
    let label = parts[parts.length - 1];

    // Replace underscores/dashes with spaces
    label = label.replace(/[_\-]/g, " ");

    // Insert a space before capital letters (for camelCase or PascalCase)
    label = label.replace(/([a-z])([A-Z])/g, "$1 $2");

    // Capitalize each word
    label = label.replace(/\b\w/g, (c) => c.toUpperCase());

    return label.trim();
  }

  // function formatKeyLabel(key: string): string {
  //   // Remove known prefixes like attr. or custom_attr.
  //   key = key.replace(/^(attr\.|custom_attr\.)/, "");

  //   const parts = key.split(".");
  //   let label = parts[parts.length - 1];

  //   // Replace underscores/dashes with spaces and capitalize each word
  //   label = label.replace(/[_\-]/g, " ");
  //   label = label.replace(/\b\w/g, (c) => c.toUpperCase());

  //   return label;
  // }

  // ---------------------------
  // Insert one field at selection in Word
  // ---------------------------
  const insertFieldAtSelection = async (key: string, value: string) => {
    try {
      await Word.run(async (context) => {
        let range = context.document.getSelection();

        // ✅ Handle image URLs
        if (
          /^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(value) ||
          value.includes("supabase.co/storage/")
        ) {
          const response = await fetch(value);
          const blob = await response.blob();
          const base64 = await blobToBase64(blob);

          range.insertInlinePictureFromBase64(base64, "Replace");
          await context.sync();

          setMessage(`🖼️ Inserted image from ${key}`);
          return;
        }

        // ✅ Handle text content
        const paragraphs = value
          .replace(/\r/g, "")
          .split(/\n{2,}/)
          .filter((p) => p.trim() !== "");

        // Clear selection before inserting
        range.insertText("", "Replace");

        for (let i = 0; i < paragraphs.length; i++) {
          const para = paragraphs[i].trim();

          if (/^[-•]\s/.test(para)) {
            // bullet list
            const lines = para
              .split(/\n|(?=-\s)/)
              .map((line) => line.replace(/^[-•]\s*/, "").trim())
              .filter((line) => line.length > 0);

            for (const line of lines) {
              const p = range.insertParagraph("", "After");
              insertFormattedText(p, line);
              p.style = "List Paragraph";
              range = p.getRange("End");
            }
          } else {
            const p = range.insertParagraph("", "After");
            insertFormattedText(p, para);
            range = p.getRange("End");
          }

          if (i < paragraphs.length - 1) {
            range.insertParagraph("", "After");
          }
        }

        await context.sync();
      });

      setMessage(`✅ Inserted "${key}" value into document`);
    } catch (err) {
      console.error(err);
      setError("⚠️ Failed to insert field into Word document");
    }
  };

  // ---------------------------
  // Helper: convert Blob → Base64
  // ---------------------------
  function blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64data = (reader.result as string).split(",")[1];
        resolve(base64data);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  // ---------------------------
  // Helper: Bold Markdown (**text**)
  // ---------------------------
  function insertFormattedText(paragraph: Word.Paragraph, text: string) {
    const boldRegex = /\*\*(.*?)\*\*/g;
    let match;
    let lastIndex = 0;

    while ((match = boldRegex.exec(text)) !== null) {
      const beforeText = text.substring(lastIndex, match.index);
      if (beforeText) paragraph.insertText(beforeText, "End");

      const boldText = match[1];
      const boldRange = paragraph.insertText(boldText, "End");
      boldRange.font.bold = true;

      lastIndex = match.index + match[0].length;
    }

    const remaining = text.substring(lastIndex);
    if (remaining) paragraph.insertText(remaining, "End");
  }

  // ---------------------------
  // UI Rendering
  // ---------------------------
  return (
    <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
      <AppHeader />

      {/* Loading overlay */}
      {loading && (
        <Box
          position="fixed"
          top={0}
          left={0}
          right={0}
          bottom={0}
          display="flex"
          alignItems="center"
          justifyContent="center"
          bgcolor="rgba(255,255,255,0.7)"
          zIndex={9999}
        >
          <CircularProgress size={60} />
        </Box>
      )}

      {/* Snackbar Alerts */}
      <Snackbar
        open={!!message || !!error}
        autoHideDuration={4000}
        onClose={() => {
          setMessage(null);
          setError(null);
        }}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        {message ? (
          <Alert severity="success" sx={{ width: "100%" }}>
            {message}
          </Alert>
        ) : error ? (
          <Alert severity="error" sx={{ width: "100%" }}>
            {error}
          </Alert>
        ) : null}
      </Snackbar>

      <Box display="flex" flexGrow={1}>
        {/* Sidebar */}
        <Box
          p={3}
          width={340}
          bgcolor="#fff"
          borderRight="1px solid #ddd"
          display="flex"
          flexDirection="column"
          sx={{
            boxShadow: "0px -2px 8px rgba(0,0,0,0.06), 2px 0 6px rgba(0,0,0,0.08)",
            borderTop: "1px solid #eee",
            position: "relative",
            zIndex: 2,
          }}
        >
          {/* <Divider sx={{ mb: 2 }} /> */}

          {/* List of Inspections */}
          <Box flexGrow={1} overflow="auto" pr={1}>
            {inspections.length > 0 ? (
              inspections.map((insp) => (
                <Card
                  key={insp.id}
                  // sx={{
                  //   mb: 2,
                  //   borderRadius: 2,
                  //   boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
                  //   transition: "0.3s",
                  //   "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
                  // }}

                  sx={{
                    marginTop: "9px",

                    // mb: 2.5,
                    borderRadius: 3,
                    backgroundColor: "#ffffff",
                    border: "1px solid #e3eaf7",
                    transition: "all 0.25s ease",
                    boxShadow:
                      "0 2px 5px rgba(0,0,0,0.05), 0 1px 3px rgba(0,0,0,0.03)",
                    "&:hover": {
                      transform: "translateY(-3px)",
                      boxShadow:
                        "0 6px 16px rgba(0,0,0,0.08), 0 3px 8px rgba(0,0,0,0.05)",
                      borderColor: "#0078d4",
                    },
                  }}
                >
                  <CardContent>
                    <Typography variant="subtitle1" fontWeight={600}
                      color="#1e3a8a">
                      {insp.attr?.material || "Unknown Material"}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
                    </Typography>
                  </CardContent>

                  <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
                    <Button
                      variant="outlined"
                      size="small"
                      onClick={() => setSelectedInspection(insp)}
                    >
                      Preview
                    </Button>
                  </CardActions>
                </Card>
              ))
            ) : (
              !loading && <Typography>No inspections found.</Typography>
            )}
          </Box>

          <Button variant="contained" onClick={getReports}
            sx={{
              mt: "auto",
              borderRadius: 2,
              textTransform: "none",
              fontWeight: 600,
              backgroundColor: "#0078d4",
              "&:hover": { backgroundColor: "#005ea2" },
            }}
          >
            Refresh Data
          </Button>
        </Box>

        {/* Modal for Preview */}
       <Modal
  open={!!selectedInspection}
  onClose={() => setSelectedInspection(null)}
  sx={{
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    p: 2,
    bgcolor: "rgba(0, 0, 0, 0.25)",
    backdropFilter: "blur(3px)",
  }}
>
  <Paper
    sx={{
      position: "relative",
      width: { xs: "100%", sm: "90%", md: 700 },
      maxHeight: "86vh",
      overflowY: "auto",
      p: 2,
      borderRadius: 3,
      backgroundColor: "#f9fbff",
      border: "1px solid #e3eaf7",
      boxShadow:
        "0 8px 24px rgba(0,0,0,0.15), 0 3px 10px rgba(0,0,0,0.05)",
      transition: "all 0.3s ease",
    }}
  >
    <IconButton
      onClick={() => setSelectedInspection(null)}
      sx={{
        position: "absolute",
        top: 12,
        right: 12,
        color: "grey.600",
        "&:hover": { color: "#0078d4" },
      }}
    >
      <CloseIcon />
    </IconButton>

    {selectedInspection && (
      <>
        <Typography
          variant="h6"
          gutterBottom
          sx={{ color: "#1e3a8a", fontWeight: 700 }}
        >
          Reports
        </Typography>
        <Divider sx={{ mb: 2, borderColor: "#d4dff0" }} />

        <Stack spacing={1.5}>
          {Object.entries(flattenInspection(selectedInspection))
            .filter(([key]) => !["id", "cid", "uid", "created_at", "attr.no"].includes(key))
            .map(([key, value]) => {
              const valStr = String(value || "");
              const isLong = valStr.length > 150;
              const isExpanded = expandedFields[key];
              const displayValue = isExpanded
                ? valStr
                : isLong
                  ? `${valStr.substring(0, 150)}...`
                  : valStr;

              return (
                <Box
                  key={key}
                  display="flex"
                  alignItems="flex-start"
                  justifyContent="space-between"
                  bgcolor="#ffffff"
                  border="1px solid #e3eaf7"
                  borderRadius={2}
                  px={1.8}
                  py={1.8}
                  boxShadow="0 1px 3px rgba(0,0,0,0.06)"
                  sx={{
                    transition: "all 0.2s ease",
                    "&:hover": {
                      boxShadow: "0 3px 8px rgba(0,0,0,0.1)",
                      borderColor: "#0078d4",
                    },
                  }}
                >
                  <Box flex={1} mr={1}>
                    <Typography
                      variant="subtitle2"
                      fontWeight={600}
                      sx={{ color: "#1e3a8a", mb: 0.5 }}
                    >
                      {formatKeyLabel(key)}
                    </Typography>

                    {/* Show image preview if URL is image */}
                    {/^https?:\/\/.*\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(valStr) ||
                    valStr.includes("supabase.co/storage/") ? (
                      <Box mt={1}>
                        <img
                          src={valStr}
                          alt={key}
                          style={{
                            maxWidth: "100%",
                            maxHeight: "150px",
                            borderRadius: "8px",
                            border: "1px solid #e0e6f1",
                          }}
                        />
                      </Box>
                    ) : (
                      <Typography
                        variant="body2"
                        color="text.secondary"
                        sx={{ whiteSpace: "pre-wrap", lineHeight: 1.5 }}
                      >
                        {displayValue}
                      </Typography>
                    )}

                    {isLong && (
                      <Button
                        size="small"
                        variant="text"
                        onClick={() =>
                          setExpandedFields((prev) => ({
                            ...prev,
                            [key]: !prev[key],
                          }))
                        }
                        sx={{
                          mt: 0.5,
                          textTransform: "none",
                          color: "#0078d4",
                          fontWeight: 500,
                        }}
                      >
                        {isExpanded ? "Show less" : "Show more"}
                      </Button>
                    )}
                  </Box>

                  {/* ✅ Insert button beside each field */}
                  <IconButton
                    color="primary"
                    onClick={() => insertFieldAtSelection(key, valStr)}
                    title="Insert this field into Word"
                    sx={{
                      bgcolor: "#eef4ff",
                      "&:hover": { bgcolor: "#d9e8ff" },
                    }}
                  >
                    <AddIcon />
                  </IconButton>
                </Box>
              );
            })}
        </Stack>
      </>
    )}
  </Paper>
</Modal>

      </Box>
    </Box>
  );
}
