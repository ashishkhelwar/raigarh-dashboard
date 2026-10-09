import React from "react";
import { createRoot } from "react-dom/client";
import DivisionDashboard from "./DivisionDashboard.jsx";
import "./index.css";

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <DivisionDashboard />
  </React.StrictMode>
);
