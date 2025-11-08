import React from "react";
import { Box, Typography, Button } from "@mui/material";
import { insertIntoWord } from "../utils/officeHelpers";

interface Inspection {
  id: string;
  client: string;
  site: string;
  date: string;
  findings: string;
}

interface Props {
  inspection: Inspection;
  onBack: () => void;
}

export default function InspectionPreview({ inspection, onBack }: Props) {
  const handleInsert = async () => {
    await insertIntoWord(inspection);
  };

  return (
    <Box sx={{ p: 2 }}>
      <Typography variant="h6">Inspection Summary</Typography>
      <Typography><b>Client:</b> {inspection.client}</Typography>
      <Typography><b>Site:</b> {inspection.site}</Typography>
      <Typography><b>Date:</b> {inspection.date}</Typography>
      <Typography><b>Findings:</b> {inspection.findings}</Typography>

      <Button variant="outlined" sx={{ mt: 2 }} onClick={onBack}>Back</Button>
      <Button variant="contained" sx={{ mt: 2, ml: 1 }} onClick={handleInsert}>
        Insert into Word
      </Button>
    </Box>
  );
}
