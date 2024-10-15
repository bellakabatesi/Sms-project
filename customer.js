const fs = require('fs');
const XLSX = require('xlsx');
//const mammoth = require('mammoth');
const axios = require('axios');

// File paths
var excelFilePath = '/home/bella/Documents/Customer_PhoneNumber.xlsx';
//const wordFilePath = '/home/bella/Documents/message.docx';

// Read the Excel file
const readExcel = () => {
  const workbook = XLSX.readFile(excelFilePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Get the first sheet
  const data = XLSX.utils.sheet_to_json(sheet);

  // Log the data to verify structure
  console.log(data.length);

  // Adjust to map the columns based on your Excel file's structure
  return data.map(customer => ({
    phone: customer['ResPhone1']  // Assuming the column for phone numbers is 'Phone Number'
  }));
};

// Read the Word document
// const readWordDoc = () => {
//   return new Promise((resolve, reject) => {
//     fs.readFile(wordFilePath, (err, data) => {
//       if (err) reject(err);
//       mammoth.extractRawText({ buffer: data })
//         .then(result => resolve(result.value))
//         .catch(reject);
//     });
//   });
// };

// Function to send message
const sendMessage = (customer, message) => {
  let data = JSON.stringify({
    "ohereza": "MUGANGA",
    "ubutumwa": "Munyamuryango,"+"\n"+"mu cyumweru cyahariwe serivisi zo kwita ku bakiliya, twishimira kubakira neza no gukorana kugirango mugere ku ntego zanyu."+"\n"+"Mugire icyumweru cyiza!",
    "msgid": customer+"week",
    "kuri": customer.phone,
    "client": "mugangasacco",
    "password": "3y2g4y9w3b1r",
    "receivedlr": "0",
    "messagetype": "binary"
  });
  
  let config = {
    method: 'post',
    maxBodyLength: Infinity,
    url: 'https://api.sms.rw/',
    headers: { 
      'Content-Type': 'application/json'
    },
    data : data
  };
  
  axios.request(config)
  .then((response) => {
    console.log(JSON.stringify(response.data));
  })
  .catch((error) => {
    console.log(error);
  });
};

const main = async () => {
  try {
    const customers = readExcel(); // Get customer list from Excel
    //const message = await readWordDoc(); // Get the message from Word document

    customers.forEach((customer) => {
      sendMessage(customer); // Send message to each customer
    });

  } catch (error) {
    console.error('Error:', error);
  }
};

main();
