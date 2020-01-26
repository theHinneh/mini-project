/** @format */

const express = require('express');
const router = express.Router();

router.get('/file_renamer', (req, res, next) => {
  res.setHeader('Content-Type', 'text/html');
  res.status(301).redirect('http://localhost:8000/');
});

module.exports = router;
