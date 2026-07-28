import React from "react";
import ReactDOM from "react-dom/client";
import ShipmentApp from "./ShipmentApp";

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("/sw.js"));
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <ShipmentApp />
  </React.StrictMode>
);
