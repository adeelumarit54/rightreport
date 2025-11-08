import { Typography } from "@mui/material";
import { motion, AnimatePresence } from "framer-motion";
import React, { useEffect, useState } from "react";
import WifiOffIcon from '@mui/icons-material/WifiOff';

export default function NetworkOverlay() {
  const [isOnline, setIsOnline] = useState(navigator.onLine);

  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  return (
    <AnimatePresence>
      {!isOnline && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.4 }}
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0, 0, 0, 0.7)",
            zIndex: 9999,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexDirection: "column",
            color: "#fff",
            backdropFilter: "blur(3px)",
          }}
        >
          <WifiOffIcon style={{ fontSize: 60, marginBottom: 16 }} />
          <Typography variant="h6">⚠️ You are offline</Typography>
          <Typography>Please check your internet connection.</Typography>
        </motion.div>
      )}
    </AnimatePresence>
  );
}
