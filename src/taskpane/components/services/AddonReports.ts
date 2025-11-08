// src/hooks/fetchAddonReports.ts
// export interface Inspection {
//   id: string;
//   attr?: {
//     no?: string;
//     material?: string;
//     location?: string;
//     inspector?: string;
//     date?: string;
//   };
//   ai_notes?: {
//     inspectorObservations?: string;
//   };
//   notes?: {
//     inspectorObservations?: string;
//   };
// }

// src/taskpane/services/AddonReports.ts
export interface Inspection {
  audio: any;
  signature: any;
  custom_attr: any;
  id: string;
  attr?: {
    no?: string;
    material?: string;
    location?: string;
    inspector?: string;
    date?: string;
    client?: string;   // ✅ Add this line
  };
  ai_notes?: {
    referenceDocuments: any;
    nonConformances: any;
    recommendations: any;
    inspectorObservations?: string;
  };
  notes?: {
    inspectionDetails: any;
    referenceDocuments: any;
    inspectorObservations?: string;
  };
}


export const fetchAddonReports = async (): Promise<Inspection[]> => {
  const token = localStorage.getItem("token");
  const refreshToken = localStorage.getItem("refresh_token");

  if (!token) {
    console.warn("No token found in localStorage");
    throw new Error("No token found");
  }

  const headers = new Headers();
  headers.append("Content-Type", "application/json");
  headers.append("Authorization", `Bearer ${token}`);
  if (refreshToken) headers.append("x-refresh-token", refreshToken);

  const response = await fetch("https://app.right-report.com/api/addon-reports", {
    method: "POST",
    headers,
  });

  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }

  const result = await response.json();
  return result?.data || [];
};
