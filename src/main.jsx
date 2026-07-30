import React from "react";
import ReactDOM from "react-dom/client";
import "./storage.js"; // installs window.storage before App.jsx uses it
import App from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
