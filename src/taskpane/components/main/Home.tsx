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
import AppHeader from "./AppHeader";
import { fetchAddonReports, Inspection } from "../services/AddonReports";
import stringSimilarity from "string-similarity";

export default function Home() {
  const [inspections, setInspections] = useState<Inspection[]>([]);
  const [selectedInspection, setSelectedInspection] = useState<Inspection | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [message, setMessage] = useState<string | null>(null);

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
  // Client-specific section titles
  // ---------------------------
  const templatesByClient: Record<string, string[]> = {
    WoodTech: [
      "Introduction / Scope",
      "Purpose of Inspection",
      "Reference Documents and Standards",
      "Findings",
      "Photos",
    ],
    SteelWorks: ["Inspection Overview", "Observations", "Defects", "Conclusion"],
  };

  // ---------------------------
  // Insert into Word
  // ---------------------------
  // const insertIntoWord = async (inspection: Inspection) => {
  //   const clientName = inspection.attr?.client || "WoodTech";
  //   const clientSections = templatesByClient[clientName] || templatesByClient["WoodTech"];

  //   const reportSections = ["Purpose of Inspection", "Findings", "Recommendations", "Photos"];

  //   const matches = reportSections.map((section) => {
  //     const { bestMatch } = stringSimilarity.findBestMatch(section, clientSections);
  //     return {
  //       reportSection: section,
  //       matchedTo: bestMatch.target,
  //     };
  //   });

  //   try {
  //     await Word.run(async (context) => {
  //       const paragraphs = context.document.body.paragraphs;
  //       paragraphs.load("items");
  //       await context.sync();

  //       for (const match of matches) {
  //         const paragraph = paragraphs.items.find((p) =>
  //           p.text.toLowerCase().includes(match.matchedTo.toLowerCase())
  //         );

  //         if (paragraph) {
  //           paragraph.insertParagraph(
  //             `${match.reportSection}: ${
  //               inspection.ai_notes?.inspectorObservations || "No data available"
  //             }`,
  //             "After"
  //           );
  //         }
  //       }

  //       await context.sync();
  //     });

  //     setMessage(`✅ Report inserted successfully for ${cliyentName}!`);
  //   } catch (error) {
  //     console.error("Error inserting into Word:", error);
  //     setError("⚠️ Failed to insert report into Word.");
  //   }
  // };



const getselectedInspection = (inspection:Inspection) => {
console.log("Selected Inspection:", inspection);
}


  const insertIntoWord = async (inspection: Inspection) => {



  try {
    // 1️⃣ Flatten all possible key/value pairs from inspection
    const fields: Record<string, string> = {};

    const extractFields = (obj: any, prefix = "") => {
      if (!obj || typeof obj !== "object") return;
      for (const [key, value] of Object.entries(obj)) {
        const fullKey = prefix ? `${prefix}.${key}` : key;
        if (typeof value === "object") extractFields(value, fullKey);
        else if (typeof value === "string" && value.trim() !== "")
          fields[key] = value;
      }
    };

    extractFields(inspection);
    extractFields(inspection.attr);
    extractFields(inspection.notes);
    extractFields(inspection.ai_notes);
    extractFields(inspection.custom_attr);

    // 2️⃣ Prepare list of field names to match in Word
    const fieldKeys = Object.keys(fields);
    console.log("Detected report fields:", fieldKeys);

    // 3️⃣ Word interaction
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const paraItems = paragraphs.items.map((p) => p.text.toLowerCase());

      for (const key of fieldKeys) {
        // find the best matching paragraph in Word
        const { bestMatch } = stringSimilarity.findBestMatch(
          key.toLowerCase(),
          paraItems
        );
        const bestIndex = paraItems.findIndex(
          (p) => p === bestMatch.target
        );

        if (bestIndex >= 0 && bestMatch.rating > 0.5) {
          const paragraph = paragraphs.items[bestIndex];
          const value = fields[key];
          paragraph.insertParagraph(`${key}: ${value}`, "After");
        }
      }

      await context.sync();
    });

    setMessage("✅ Dynamic report inserted successfully!");
  } catch (error) {
    console.error("Error inserting into Word:", error);
    setError("⚠️ Failed to insert dynamic report into Word.");
  }
};

// async function insertIntoWord(inspection: any) {
//   await Word.run(async (context) => {
//     const body = context.document.body;

//     // Extract key values safely
//     const obs =
//       inspection.inspectorObservations?.trim() ||
//       inspection.transcript?.trim() ||
//       "No observations available";

//     const refDocs =
//       inspection.referenceDocuments?.trim() || "No reference documents available";

//     const mdb =
//       inspection.custom_attr?.["MDB Completion"] ||
//       "No MDB Completion status provided";

//     // Map of placeholders to values
//     const replacements: Record<string, string> = {
//       inspectorObservations: obs,
//       referenceDocuments: refDocs,
//       "MDB Completion": mdb,
//       Summary:
//         "Overall, the inspection has been completed. Please review observations for non-conformances.",
//     };

//     // Loop and replace placeholders
//     for (const [key, value] of Object.entries(replacements)) {
//       const searchResults = body.search("(Content from API will be inserted here)", {
//         matchCase: false,
//         matchWholeWord: false,
//       });
//       searchResults.load("items");
//       await context.sync();

//       for (const range of searchResults.items) {
//         const prev = range.paragraphs.getFirst().getPrevious();
//         prev.load("text");
//         await context.sync();

//         if (prev.text.includes(key)) {
//           range.insertText(String(value), "Replace");
//         }
//       }
//     }

//     // // Add signature (optional)
//     // if (inspection.signature?.publicUrl) {
//     //   // const base64 = await fetchAsBase64(inspection.signature.publicUrl);
//     //   body.insertParagraph("Inspector Signature:", "End");
//     //   // body.insertInlinePictureFromBase64(base64, "End");
//     // }

//     body.insertParagraph("\n--- End of Report ---\n", "End");

//     await context.sync();
//   });
// }

  

// Helper: Flatten inspection object (nested keys)
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

// Helper: Insert only one selected field
const insertSingleField = async (key: string, value: string) => {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(`${key}: ${value}`, "Replace");
      await context.sync();
    });
    setMessage(`✅ Inserted "${key}" into Word`);
  } catch (error) {
    console.error("Error inserting field:", error);
    setError("⚠️ Failed to insert field into Word.");
  }
};

  return (
    <Box display="flex" flexDirection="column" height="100vh" bgcolor="#f4f6f8">
      <AppHeader />

      {/* Loading Overlay */}
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
          sx={{ boxShadow: "2px 0 4px rgba(0,0,0,0.05)" }}
        >
         
          <Divider sx={{ mb: 2 }} />

          {/* List of Inspections */}
          <Box flexGrow={1} overflow="auto" pr={1}>
            {inspections.length > 0 ? (
              inspections.map((insp) => (
                <Card
                  key={insp.id}
                  sx={{
                    mb: 2,
                    borderRadius: 2,
                    boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
                    transition: "0.3s",
                    "&:hover": { boxShadow: "0 4px 10px rgba(0,0,0,0.15)" },
                  }}
                >
                  <CardContent>
                    <Typography variant="subtitle1" fontWeight={600}>
                      #{insp.attr?.no || "N/A"} – {insp.attr?.material || "Unknown Material"}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      {insp.attr?.location || "Unknown"} • {insp.attr?.inspector || "No Inspector"}
                    </Typography>
                  </CardContent>

                  <CardActions sx={{ justifyContent: "flex-end", px: 2, pb: 2 }}>
                    <Button
                      variant="outlined"
                      size="small"
                      onClick={() => getselectedInspection(insp)}
                    >
                      Preview
                    </Button>
                    <Button
                      variant="contained"
                      size="small"
                      onClick={() => insertIntoWord(insp)}
                    >
                      Insert
                    </Button>
                  </CardActions>
                </Card>
              ))
            ) : (
              !loading && <Typography>No inspections found.</Typography>
            )}
          </Box>

          <Button variant="contained" onClick={getReports} sx={{ mt: "auto" }}>
            Refresh Data
          </Button>
        </Box>

        {/* Modal for Preview */}
       {/* Modal for Preview */}
<Modal
  open={!!selectedInspection}
  onClose={() => setSelectedInspection(null)}
  sx={{
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    p: 2,
  }}
>
  <Paper
    sx={{
      position: "relative",
      width: { xs: "100%", sm: "90%", md: 700 },
      maxHeight: "86vh",
      overflowY: "auto",
      p: 3,
      borderRadius: 3,
      boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
    }}
  >
    {/* Close Button */}
    <IconButton
      onClick={() => setSelectedInspection(null)}
      sx={{ position: "absolute", top: 8, right: 8, color: "grey.600" }}
    >
      <CloseIcon />
    </IconButton>

    {selectedInspection && (
      <>
        <Typography variant="h6" gutterBottom>
          Inspection Report #{selectedInspection.attr?.no || "N/A"}
        </Typography>
        <Divider sx={{ mb: 2 }} />

        {/* Flatten all fields dynamically */}
        <Stack spacing={1}>
          {Object.entries(flattenInspection(selectedInspection)).map(([key, value]) => (
            <Box
              key={key}
              display="flex"
              alignItems="center"
              justifyContent="space-between"
              bgcolor="#f9f9f9"
              borderRadius={1}
              px={2}
              py={1}
            >
              <Box flex={1}>
                <Typography variant="subtitle2" fontWeight={600}>
                  {key}
                </Typography>
                <Typography
                  variant="body2"
                  color="text.secondary"
                  sx={{ whiteSpace: "pre-wrap" }}
                >
                  {String(value)}
                </Typography>
              </Box>

              {/* Insert icon button */}
              <IconButton
                color="primary"
                onClick={() => insertSingleField(key, String(value))}
              >
                <i className="ms-Icon ms-Icon--Add" style={{ fontSize: 20 }} />
              </IconButton>
            </Box>
          ))}
        </Stack>
      </>
    )}
  </Paper>
</Modal>

      </Box>
    </Box>
  );
}
