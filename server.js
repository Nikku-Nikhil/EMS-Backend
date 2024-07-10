const express = require("express");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const nodemailer = require("nodemailer");
const QRCode = require("qrcode");
const mongoose = require("mongoose");
const dotenv = require("dotenv").config();
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const cors = require("cors"); // Import the cors package

const connectDb = require("./config/dbConnection");

const app = express();
const port = process.env.PORT || 3000;

connectDb();

app.use(bodyParser.json());
app.use(cors()); // Use the cors middleware

// Configure multer for file uploads
const upload = multer({ dest: "uploads/" });

// MongoDB Schema and Model
const studentSchema = new mongoose.Schema({
  name: String,
  email: String,
  admissionId: String,
  phoneNumber: String,
  isApproved: { type: Boolean, default: false },
});

const Student = mongoose.model("Student", studentSchema);

const fileSchema = new mongoose.Schema({
  filename: String,
  contentType: String,
  data: Buffer,
});

const File = mongoose.model("File", fileSchema);

// Setup Nodemailer
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Function to process each student and send QR code
const processStudent = async (student) => {
  const { name, email, admissionId, phoneNumber } = student;

  const qrCodeData = `${
    process.env.BASE_URL
  }/scanQrCode?name=${encodeURIComponent(name)}&email=${encodeURIComponent(
    email
  )}&admissionId=${encodeURIComponent(
    admissionId
  )}&phoneNumber=${encodeURIComponent(phoneNumber)}`;

  try {
    const qrCode = await QRCode.toDataURL(qrCodeData);

    // Send email with QR code
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: email,
      subject: "Your QR Code",
      html: `Hi ${name},<br><br>

Please find your QR code attached below.<br><br>

<strong>Note: Do not scan this QR code by yourself, as it is for one-time use.</strong><br><br>

Thank you.`,
      attachments: [
        {
          filename: "qrcode.png",
          path: qrCode,
        },
      ],
    };

    await transporter.sendMail(mailOptions);

    // Save to MongoDB
    const newStudent = new Student({
      name,
      email,
      admissionId,
      phoneNumber,
    });

    await newStudent.save();

    console.log(`Email sent and data saved for ${name}`);
  } catch (error) {
    console.error(`Error processing student ${name}:`, error);
  }
};

// Endpoint to upload the Excel file
app.post("/uploadExcel", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded");
  }

  const file = new File({
    filename: req.file.originalname,
    contentType: req.file.mimetype,
    data: fs.readFileSync(req.file.path),
  });

  try {
    const savedFile = await file.save();
    fs.unlinkSync(req.file.path); // Remove the file from the server after saving to DB
    res.send({
      message: "File uploaded and saved successfully",
      fileId: savedFile._id,
    });
  } catch (error) {
    console.error("Error saving file:", error);
    res.status(500).send("Internal Server Error");
  }
});

// Endpoint to trigger the process
app.get("/sendQrCodes", async (req, res) => {
  const fileId = req.query.fileId;

  if (!fileId) {
    return res.status(400).send("File ID is required");
  }

  try {
    const file = await File.findById(fileId);

    if (!file) {
      return res.status(404).send("File not found");
    }

    // Load the workbook from the file data
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.data);

    const worksheet = workbook.getWorksheet(1);
    const studentsData = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber !== 1) {
        // Skip header row
        const student = {
          name: row.getCell(1).value,
          email: row.getCell(2).value,
          admissionId: row.getCell(3).value,
          phoneNumber: row.getCell(4).value,
        };
        studentsData.push(student);
      }
    });

    for (const student of studentsData) {
      await processStudent(student);
    }

    // Delete the file from the database after processing
    await File.findByIdAndDelete(fileId);

    res.send("QR codes sent, data saved, and file deleted successfully");
  } catch (error) {
    console.error("Error processing students:", error);
    res.status(500).send("Internal Server Error");
  }
});

// Endpoint to handle QR code scanning
app.get("/scanQrCode", async (req, res) => {
  const { name, email, admissionId, phoneNumber } = req.query;

  if (!name || !email || !admissionId || !phoneNumber) {
    return res
      .status(400)
      .send("Name, email, admissionId, and phoneNumber are required");
  }

  try {
    const student = await Student.findOne({
      name,
      email,
      admissionId,
      phoneNumber,
    });

    if (!student) {
      return res.status(404).send(`
       <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
           text-align: center;
            margin-top: 50px;
          }
          .message {
            color: red;
             font-size: 8rem;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="message">Student not Found</div>
      </body>
    </html>`);
    }

     if (student.isApproved) {
      return res.status(400).send(`
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
             text-align: center;
            margin-top: 50px;
          }
          .message {
            color: red;
            font-size: 8rem;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="message">QR code already scanned</div>
      </body>
    </html>
  `);
    }

    student.isApproved = true;
    await student.save();

    res.send(`
  <html>
    <head>
      <style>
        body {
          font-family: Arial, sans-serif;
          text-align: center;
          margin-top: 50px;
        }
        .message {
          color: green;
          font-size: 8rem;
          font-weight: bold;
        }
      </style>
    </head>
    <body>
      <div class="message">QR code scanned successfully</div>
    </body>
  </html>
`);
  } catch (error) {
    console.error("Error scanning QR code:", error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/downloadStudents", async (req, res) => {
  try {
    const students = await Student.find();

    if (!students || students.length === 0) {
      return res.status(404).send("No students found");
    }

    // Create a workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Students");

    // Define columns
    worksheet.columns = [
      { header: "Name", key: "name", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "Admission ID", key: "admissionId", width: 15 },
      { header: "Phone Number", key: "phoneNumber", width: 15 },
      { header: "Approved", key: "isApproved", width: 10 },
    ];

    // Add rows
    students.forEach((student) => {
      worksheet.addRow({
        name: student.name,
        email: student.email,
        admissionId: student.admissionId,
        phoneNumber: student.phoneNumber,
        isApproved: student.isApproved ? "Yes" : "No",
      });
    });

    // Set response headers
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=" + "students.xlsx"
    );

    // Send workbook as response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Error downloading students:", error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/", async (req, res) => {
  res.status(200).send("Landing Page");
});

app.listen(port, () => {
  console.log(`Server running at 127.0.0.1:${port}`);
});
