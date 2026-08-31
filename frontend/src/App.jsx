import { useEffect, useState } from "react";
import "./App.css";

function App() {
  const [apiStatus, setApiStatus] = useState("Проверяем API...");
  const [dbStatus, setDbStatus] = useState("Проверяем PostgreSQL...");

  useEffect(() => {
    fetch("http://localhost:3001/api/status")
      .then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        return response.json();
      })
      .then((data) => {
        setApiStatus(data.message);
      })
      .catch((error) => {
        console.error(error);
        setApiStatus("Ошибка соединения с API");
      });

    fetch("http://localhost:3001/api/db/status")
      .then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        return response.json();
      })
      .then((data) => {
        setDbStatus(
          `PostgreSQL подключен: ${data.database}, порт ${data.port}`,
        );
      })
      .catch((error) => {
        console.error(error);
        setDbStatus("Ошибка соединения с PostgreSQL");
      });
  }, []);

  return (
    <main>
      <h1>Price Control V3</h1>

      <p>{apiStatus}</p>
      <p>{dbStatus}</p>
    </main>
  );
}

export default App;
