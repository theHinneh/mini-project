/** @format */

const http = require('http');
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');

const d2_wswsRoute = require('./routes/d2_wsws');
const gillete_appRoute = require('./routes/gillete_app');
const file_renamerRoute = require('./routes/file_renamer');
const homeRoute = require('./routes/home');

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

app.use(d2_wswsRoute);
app.use(gillete_appRoute);
app.use(file_renamerRoute);
app.use(bodyParser.urlencoded({ extended: false }));
app.use(homeRoute);
app.use((req, res, next) => {
  res.status(404).send('<h1>Page Not Founs </h1>');
});

const server = http.createServer(app);
server.listen(3000);
