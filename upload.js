const express = require("express");
const router = express.Router();
const multer = require("multer");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

router.post("/upload", upload.single("file"), async (req, res) => {
  const { subject, content } = req.body;
  const filePath = req.file.path;
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const emailColumn = "A";
  const emails = Object.keys(sheet)
    .filter((key) => key.startsWith(emailColumn))
    .map((key) => sheet[key].v);

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: "gomathy2511@gmail.com",
      pass: process.env.EMAIL_PASSWORD,
    },
  });

  const mailOptions = {
    from: "gomathy2511@gmail.com",
    to: emails.join(","),
    subject: subject,
    text: content,
  };

  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.error(error);
      res.status(500).json({ message: "Error sending email" });
    } else {
      console.log("Email sent: " + info.response);
      res.json({ message: "Email sent successfully" });
    }
  });
});

module.exports = router;
