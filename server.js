const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const pdf = require('html-pdf');
const ejs = require('ejs');
const app = express();


mongoose.connect('mongodb+srv://admin:Kanishika23@cluster0.sjcpv5s.mongodb.net/?retryWrites=true&w=majority').then(() => {
  console.log("Database connected")
})
const db = mongoose.connection;

const dataSchema = new mongoose.Schema({
  name: String,
  email: { type: String, unique: true },
  mobileno:{ type: String,unique:true },
  noOfTrees: Number,
  amount: Number,
});

const Data = mongoose.model('Data', dataSchema);

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.set("view engine", "ejs");

app.get('/', (req, res) => {
  res.render('upload');
})

app.post('/upload', upload.single('file'), async (req, res) => {
  const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  // Processing data before insertion
  const processedData = jsonData.map(user => {
    const noOfTrees = Math.floor(user.amount / 100); // Calculate trees donated
    return {
      name: user.name,
      email: user.email,
      mobileno: user.mobileno,
      amount: user.amount,
      noOfTrees: user.noOfTrees
    };
  });

  try {
    await Data.insertMany(processedData);

    // Loop through the processed data to generate certificates and send emails
    processedData.forEach(async (user) => {
      // Render certificate.ejs with dynamic data
      ejs.renderFile('certificate.ejs', { user }, async (err, htmlTemplate) => {
        if (err) {
          console.error(err);
          return res.status(500).send('Error rendering EJS template');
        }

        // Generate PDF from HTML
        pdf.create(htmlTemplate).toFile(`certificate_${ user.name }.pdf`, async function (pdfErr, pdfPath) {
          if (pdfErr) {
            console.error(pdfErr);
            return res.status(500).send('Error generating PDF');
          }

          
          let transporter = nodemailer.createTransport({
           
            service: 'Gmail',
            auth: {
              user: 'kanishika0722.be21@chitkara.edu.in',
              pass: 'Nehagoyal1983'
            }
          });

          // Email options
          let mailOptions = {
            from: 'kanishika0722.be21@chitkara.edu.in',
            to: user.email,
            subject: 'Certificate of Tree Donation',
            text: 'Certificate attached.',
            attachments: [{
              filename: `certificate_${ user.name }.pdf`,
              path: pdfPath.filename

            }]
        };

        // Send email with attachment
        try {
          await transporter.sendMail(mailOptions);
          console.log(`Certificate sent to ${user.email}`);
        } catch (emailErr) {
          console.error(emailErr);
        }
      });
    });
  });

res.status(200).send('Data uploaded successfully');
  } catch (err) {
  console.error(err);
  res.status(500).send('Error uploading data');
}
});
app.listen(3003, function () {
  console.log("server connected");
})