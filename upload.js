const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");

const router = express.Router();

// Set storage for uploaded files
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

// Create upload function using multer
const upload = multer({ storage }).single("file");

router.post("/upload", upload, (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ message: "No file uploaded" });
    }
    const workbook = xlsx.readFile(file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    const emails = data.map((item) => item["Email"]);

    // Create nodemailer transporter
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: "gomathy2511@gmail.com",
        pass: process.env.EMAIL_PASSWORD,
      },
    });

    // Send emails to extracted emails
    emails.forEach((email) => {
      const mailOptions = {
        from: "gomathy2511@gmail.com",
        to: email,
        subject: "Bulk Email",
        text: "Welcome To Bulk Email Sending App",
      };
      transporter.sendMail(mailOptions, (err, info) => {
        if (err) {
          console.log(err);
        } else {
          console.log(`Email sent to ${email}: ${info.response}`);
        }
      });
    });

    res.status(200).json({ message: "Emails sent successfully" });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Server error" });
  }
});

module.exports = router;
