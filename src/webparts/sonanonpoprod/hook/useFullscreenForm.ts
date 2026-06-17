// src/hooks/useFullscreenForm.ts
import { useEffect } from "react";
import { useLocation } from "react-router-dom";


export default function useFullscreenForm() {
  const location = useLocation();

  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const fromState = (location.state as any)?.fullscreen === true;
    const fromQuery = params.get("fs") === "1";
    const shouldFullscreen = fromState || fromQuery;

    if (shouldFullscreen) {
      document.body.classList.add("no-sidebar");
    } else {
      document.body.classList.remove("no-sidebar");
    }

    return () => {
      document.body.classList.remove("no-sidebar");
    };
  }, [location]);
}