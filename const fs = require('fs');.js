const fs = require('fs');
const XLSX = require('xlsx');
const mammoth = require('mammoth');

// Read the Excel file
const readExcel = (filePath) => {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet);
  return data;
};

// Read the Word document
const readWordDoc = (filePath) => {
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, (err, data) => {
      if (err) reject(err);
      mammoth.extractRawText({ buffer: data })
        .then(result => resolve(result.value))
        .catch(reject);
    });
  });
};

// Function to send message (you can modify it based on how you want to send messages)
const sendMessage = (customer, message) => {
  // Example: sending email (you can replace this with any other method like SMS, etc.)
  console.log(`Sending message to ${customer.name} at ${customer.email}`);
  console.log(`Message: ${message}`);
  // Here you can use an email API like Nodemailer or others to send the actual message.
};

const main = async () => {
  const excelFilePath = 'path/to/customer_list.xlsx'; // Provide path to Excel file
  const wordFilePath = 'path/to/message.docx'; // Provide path to Word document

  try {
    const customers = readExcel(excelFilePath); // Get customer list from Excel
    const message = await readWordDoc(wordFilePath); // Get the message from Word document

    customers.forEach((customer) => {
      sendMessage(customer, message); // Send message to each customer
    });

  } catch (error) {
    console.error('Error:', error);
  }
};

main();
