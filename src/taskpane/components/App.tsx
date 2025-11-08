// import React from "react";
// import NetworkStatus from "./main/NetworkStatus";
// import AppRouter from "./router/AppRouter";


// export default function App() {
//   return (
//     <>
//       <NetworkStatus />
//       <AppRouter />
//     </>
//   );
// }


import React from "react";
import AppRouter from "./router/AppRouter";
import NetworkOverlay from "./networkOverlay/NetworkOverlay";

export default function App() {
  return (
    <>
      <AppRouter />
      <NetworkOverlay /> {/* Always active */}
    </>
  );
}
