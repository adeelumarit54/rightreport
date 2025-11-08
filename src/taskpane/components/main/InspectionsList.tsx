import React, { useEffect, useState } from "react";
import { List, ListItem, ListItemText, Typography, Box, Button } from "@mui/material";
// import api from "../../api/apiClient";

interface Inspection {
  id: string;
  client: string;
  site: string;
  date: string;
  findings: string;
}

interface Props {
  onSelect: (inspection: Inspection) => void;
  onLogout: () => void;
}

// export default function InspectionsList({ onSelect, onLogout }: Props) {
export default function InspectionsList({ onLogout }: Props) {
    
  const [inspections, setInspections] = useState<Inspection[]>([]);
  const [loading, setLoading] = useState(true);

  const loadInspections = async () => {
    setLoading(true);
    try {
    //   const res = await api.get("/api/inspections");
    //   setInspections(res.data || []);
    } catch {
      alert("Failed to load inspections");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadInspections();
  }, []);

  return (
    <Box sx={{ p: 2 }}>
      <Box sx={{ display: "flex", justifyContent: "space-between", mb: 2 }}>
        <Typography variant="h6">Your Inspections</Typography>
        <Button onClick={onLogout}>Logout</Button>
      </Box>

      {loading ? (
        <Typography>Loading...</Typography>
      ) : inspections.length === 0 ? (
        <Typography>No inspections found</Typography>
      ) : (
        <List>
          {/* {inspections.map((i) => (
            <ListItem key={i.id} button onClick={() => onSelect(i)}>
              <ListItemText primary={i.client} secondary={`${i.site} • ${i.date}`} />
            </ListItem>
          ))} */}

            <ListItem >
              <ListItemText  />
            </ListItem>
        </List>
      )}
      <Button onClick={loadInspections}>Refresh</Button>
    </Box>
  );
}
