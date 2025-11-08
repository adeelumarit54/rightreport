// src/pages/Dashboard.tsx
import React from "react";
import { Button, Box, Typography } from "@mui/material";

export default function Dashboard() {
  const handleLogout = () => {
    sessionStorage.removeItem("rr_token");
    window.location.href = "/";
  };

  return (
    <Box sx={{ p: 3 }}>
      <Typography variant="h5">Welcome to RightReport Dashboard</Typography>
      <Button variant="outlined" sx={{ mt: 2 }} onClick={handleLogout}>
        Logout
      </Button>
    </Box>
  );
}
