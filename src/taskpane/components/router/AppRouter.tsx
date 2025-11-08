// import React, { useEffect, useState } from "react";
// import { HashRouter, Routes, Route, Navigate } from "react-router-dom";
// import Home from "../main/Home";
// import LoginForm from "../login/LoginForm";
// import { Box, CircularProgress } from "@mui/material";

// const isTokenValid = (token: string | null): boolean => {
//   if (!token) return false;
//   try {
//     const payload = JSON.parse(atob(token.split(".")[1]));
//     const expiry = payload.exp * 1000;
//     return Date.now() < expiry;
//   } catch (e) {
//     console.error("Invalid token format:", e);
//     return false;
//   }
// };

// export default function AppRouter() {
//   const [checking, setChecking] = useState(true);
//   const [valid, setValid] = useState(false);

//   useEffect(() => {
//     const token = localStorage.getItem("token");
//     setValid(isTokenValid(token));
//     setChecking(false);
//   }, []);

//   if (checking) {
//     return (
//       <Box
//         sx={{
//           height: "100vh",
//           display: "flex",
//           justifyContent: "center",
//           alignItems: "center",
//           backgroundColor: "#f7f9fb",
//         }}
//       >
//         <CircularProgress size={50} thickness={4} />
//       </Box>
//     );
//   }
// return (
//   <HashRouter>
//     <Routes>
//       <Route
//         path="/"
//         element={valid ? <Navigate to="/home" replace /> : <LoginForm />}
//       />
//       <Route
//         path="/home"
//         element={valid ? <Home /> : <Navigate to="/" replace />}
//       />
//     </Routes>
//   </HashRouter>
// );

// }



import React, { useEffect, useState } from "react";
import { MemoryRouter, Routes, Route, Navigate, useNavigate } from "react-router-dom";
// import Home from "../main/Home";
import Home from "../main/Maintest";

import LoginForm from "../login/LoginForm";
import { Box, CircularProgress } from "@mui/material";

const isTokenValid = (token: string | null): boolean => {
  if (!token) return false;
  try {
    const payload = JSON.parse(atob(token.split(".")[1]));
    const expiry = payload.exp * 1000;
    return Date.now() < expiry;
  } catch {
    return false;
  }
};

export default function AppRouter() {
  const [checking, setChecking] = useState(true);
  const [valid, setValid] = useState(false);

  useEffect(() => {
    const token = localStorage.getItem("token");
    setValid(isTokenValid(token));
    setChecking(false);
  }, []);

  if (checking) {
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
        <CircularProgress size={50} thickness={4} />
      </Box>
    );
  }

  return (
    <MemoryRouter>
      <Routes>
        <Route
          path="/"
          element={valid ? <Navigate to="/home" replace /> : <LoginForm />}
        />
        <Route
          path="/home"
          element={valid ? <Home /> : <Navigate to="/" replace />}
        />
      </Routes>
    </MemoryRouter>
  );
}
