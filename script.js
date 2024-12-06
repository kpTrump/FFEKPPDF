const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const xlsx = require('xlsx');

// Path to the folder containing PDF files
const pdfFolderPath = 'C:\Users\konradas.peciulis\Desktop\Empty Folder';

// Read all PDF files in the directory
fs.readdir(pdfFolderPath, (err, files) => {
  if (err) {
    console.error('Error reading directory:', err);
    return;
  }

  // Filter out non-PDF files
  const pdfFiles = files.filter(file => path.extname(file).toLowerCase() === '.pdf');
  
  // Array to store the extracted data
  let extractedData = [];

  // Loop through each PDF file and parse its content
  let filePromises = pdfFiles.map(file => {
    const filePath = path.join(pdfFolderPath, file);
    
    return new Promise((resolve, reject) => {
      const dataBuffer = fs.readFileSync(filePath);
      
      pdfParse(dataBuffer).then(data => {
        const text = data.text;

        // Extract the required fields from the PDF text
        const accountNoMatch = text.match(/Account No:\s*(\S+)/);
        const ffeReserveMatch = text.match(/FF&E Reserve:\s*([0-9,\.]+)/);

        // If both fields are found, push the values to the data array
        if (accountNoMatch && ffeReserveMatch) {
          extractedData.push([accountNoMatch[1], ffeReserveMatch[1]]);
        }

        resolve(); // Resolve the promise after processing the file
      }).catch(err => {
        console.error('Error parsing PDF:', err);
        reject();
      });
    });
  });

  // Once all PDFs are processed, create and save the Excel file
  Promise.all(filePromises).then(() => {
    // Create a worksheet from the extracted data
    const ws = xlsx.utils.aoa_to_sheet([['Account No:', 'FF&E Reserve'], ...extractedData]);
    
    // Create a workbook with the worksheet
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Extracted Data');
    
    // Write the Excel file to disk
    const excelFilePath = './extracted_data.xlsx';
    xlsx.writeFile(wb, excelFilePath);
    console.log('Excel file created:', excelFilePath);
  }).catch(() => {
    console.error('Error processing PDF files.');
  });
});
