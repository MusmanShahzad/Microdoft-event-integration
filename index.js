const express = require("express");
const bodyParser = require("body-parser");
const axios = require("axios");
require("dotenv").config();
const app = express();
const port = 8080;
// parse application/x-www-form-urlencoded
app.use(bodyParser.urlencoded({ extended: false }));
// parse application/json
app.use(bodyParser.json());

app.get("/", (req, res) => {
  res.send("Welcome to the rest api");
});

app.post("/createmeeting", async (req, res) => {
  const event = req.body;
  let token = req.headers.authorization;
  try {
    return res.send(
      (
        await axios.default.post(
          "https://graph.microsoft.com/v1.0/me/events",
          event,
          {
            headers: {
              Authorization: `Bearer ${token}`,
            },
          }
        )
      ).data
    );
  } catch (err) {
    console.log(err);
    res.send(err.response.data);
  }
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
