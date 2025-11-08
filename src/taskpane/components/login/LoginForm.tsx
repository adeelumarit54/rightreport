import React, { useState, useEffect } from "react";
import {
  TextField,
  Button,
  Box,
  CircularProgress,
  Snackbar,
  Alert,
} from "@mui/material";
import { useNavigate } from "react-router-dom"; 
import { loginUser } from "../services/login";

export default function LoginForm() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [openError, setOpenError] = useState(false);
  const [openSuccess, setOpenSuccess] = useState(false);

  const navigate = useNavigate(); 

  useEffect(() => {
    if (error) setOpenError(true);
  }, [error]);

  useEffect(() => {
    if (success) setOpenSuccess(true);
  }, [success]);

  const handleSubmit = async () => {
    setError("");
    setSuccess("");
    setLoading(true);

    const result = await loginUser(email, password);
    setLoading(false);

    if (result.error) {
      setError(result.error);
    } else if (result.jwt && result.refresh_token) {
      
      localStorage.setItem("token", result.jwt);
      localStorage.setItem("refresh_token", result.refresh_token);

      setSuccess("Login successful!");
window.location.reload();
//       if (result.jwt && result.refresh_token) {
//   localStorage.setItem("token", result.jwt);
//   localStorage.setItem("refresh_token", result.refresh_token);
//   setSuccess("Login successful!");
//   setTimeout(() => navigate("/home"), 500); // small delay
// }
      
    } else {
      setError("Invalid response from server.");
    }
  };

  return (
    <Box
      sx={{
        height: "100vh",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        backgroundColor: "#f7f9fb",
      }}
    >
      <Box sx={{ width: 340, p: 4, textAlign: "center" }}>
        <Box sx={{ mb: 3 }}>
          <img
            src="https://cdn.prod.website-files.com/688c817cfc54b7ffa81477d2/6893b1f65fe9bfb915176efe_Website%20RR%20(2).svg"
            alt="RightReport Logo"
            style={{ width: "150px", height: "auto" }}
          />
        </Box>

        <TextField
          label="Email"
          fullWidth
          size="small"
          margin="dense"
          value={email}
          // sx={{marginTop:"20px"}}
          onChange={(e) => setEmail(e.target.value)}
        />

        <TextField
          label="Password"
          fullWidth
          size="small"
          margin="dense"
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />

        <Box sx={{ mt: 3, textAlign: "center" }}>
          <Button
            variant="contained"
            onClick={handleSubmit}
            disabled={loading || !email || !password}
            sx={{
              minWidth: 120,
              px: 3,
              py: 0.8,
              textTransform: "none",
              fontWeight: 500,
              borderRadius: 2,
            }}
          >
            {loading ? <CircularProgress size={22} color="inherit" /> : "Login"}
          </Button>
        </Box>
      </Box>

      {/* Error Snackbar */}
      <Snackbar
        open={openError}
        autoHideDuration={4000}
        onClose={() => setOpenError(false)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert
          onClose={() => setOpenError(false)}
          severity="error"
          sx={{ width: "100%" }}
        >
          {error}
        </Alert>
      </Snackbar>

      {/* Success Snackbar */}
      <Snackbar
        open={openSuccess}
        autoHideDuration={4000}
        onClose={() => setOpenSuccess(false)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert
          onClose={() => setOpenSuccess(false)}
          severity="success"
          sx={{ width: "100%" }}
        >
          {success}
        </Alert>
      </Snackbar>
    </Box>
  );
}
