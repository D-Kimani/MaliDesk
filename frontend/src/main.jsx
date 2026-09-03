import React from "react";
import { createRoot } from "react-dom/client";
import MaliDesk from "./MaliDesk.jsx";
import "./index.css";

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <MaliDesk />
  </React.StrictMode>
);
