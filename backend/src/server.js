const express = require("express");
const cors = require("cors");

const app = express();

const PORT = 3001;
const HOST = "0.0.0.0";

app.use(cors());
app.use(express.json());

app.get("/api/status", (req, res) => {
  res.json({
    ok: true,
    service: "Price Control V3 API",
    message: "Price Control V3 API работает",
  });
});

app.listen(PORT, HOST, () => {
  console.log(`Price Control V3 API запущен: http://localhost:${PORT}`);
});
