import { useState } from "react";

export default function useAppApi() {
  const [message, setMessage] = useState("");

  const showMessage = (text) => {
    setMessage(text);
    setTimeout(() => setMessage(""), 3500);
  };

  const authFetch = (url, options = {}) => {
    const savedToken = sessionStorage.getItem("token");

    return fetch(url, {
      ...options,
      headers: {
        "Content-Type": "application/json",
        ...(options.headers || {}),
        Authorization: "Bearer " + savedToken,
      },
    });
  };

  return {
    authFetch,
    message,
    setMessage,
    showMessage,
  };
}
