import { useEffect, useState } from "react";
import "./App.css";

function App() {
  const [status, setStatus] = useState("роверяем соединение с сервером...");

  useEffect(() => {
    fetch("http://localhost:3001/api/status")
      .then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        return response.json();
      })
      .then((data) => {
        setStatus(data.message);
      })
      .catch((error) => {
        console.error(error);
        setStatus("шибка соединения с Price Control V3 API");
      });
  }, []);

  return (
    <main>
      <h1>Price Control V3</h1>
      <p>{status}</p>
    </main>
  );
}

export default App;
