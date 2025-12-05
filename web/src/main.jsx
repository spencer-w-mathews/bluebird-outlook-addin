import React from "react";
import { createRoot } from "react-dom/client";
import BluebirdPane from "./BluebirdPane";

Office.onReady(() => {
  const root = createRoot(document.getElementById("root"));
  root.render(<BluebirdPane />);
});
