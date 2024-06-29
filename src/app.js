const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const ExcelJS = require('exceljs');
const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();


const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;



// Middleware to handle form data

const max_upload_size_mb = 128;


const upload = multer({ 
  dest: 'uploads/',
limits: { fileSize: max_upload_size_mb * 1024 * 1024 } 
 });


function log(message) {
  const now = new Date();
  const timestamp = now.toISOString();
  console.log(`[${timestamp}] ${message}`);
}
function err(message) {
  const now = new Date();
  const timestamp = now.toISOString();
  console.error(`[${timestamp}] ${message}`);
}

// Serve static files
app.use(express.static(path.join(__dirname, '../public')));
app.use(bodyParser.urlencoded({ extended: true }));


app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(413).send('File size exceeds the limit.');
  }
  next(err);
});


// Route to handle file upload
app.post('/upload', upload.single('excelFile'), async (req, res) => {

  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded.');
    }

    const fileTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
  ];

  if (!fileTypes.includes(req.file.mimetype)) { 
      fs.unlink(req.file.path, (error) => {
          if (error) {
              err(`Error deleting file: ${error}`);
          }
          log('Uploaded file deleted successfully.');
      });
      return res.status(400).send('File is not an Excel file!');
  }

    log("uploading " + req.file.filename + " size = " + req.file.size);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);


    //https://gist.github.com/davidhq/0afed70985842cac6fc4e00f88a71bd2
    const worksheet = workbook.getWorksheet();
    if (!worksheet){
      err("null worksheet");
      throw new err('Worksheet not found in uploaded workbook');
    } 

    function getCellResult(worksheet, cellLabel) {
      if (worksheet.getCell(cellLabel).formula) {
        return parser.parse(worksheet.getCell(cellLabel).formula).result;
      } else {
        return worksheet.getCell(cellCoord.label).value;
      }
    }
    /*
            parser.on('callCellValue', function(cellCoord, done) {
                if (worksheet.getCell(cellCoord.label).formula) {
                  done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
                } else {
                  done(worksheet.getCell(cellCoord.label).value);
                }
              });
    
              parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
                var fragment = [];
            
                for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
                  var colFragment = [];
            
                  for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
                    colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
                  }
            
                  fragment.push(colFragment);
                }
            
                if (fragment) {
                  done(fragment);
                }
              });*/
    
    log(`A1 = ${worksheet.getCell("A1").value}`);

    // Clean up uploaded file
    log("deleting original file " + req.file.path);
   // fs.unlinkSync(req.file.path);
    fs.unlink(req.file.path, (error) => {
    if (error) {
      err('Error unlinking file: ' + error);
    } else {
      log('upload returned and deleted successfully. \n');
    }
  });


    if(0){
      var csv_file = req.file.path + ".csv"; 
      log("writing to " + csv_file);
      workbook.csv.writeFile(csv_file); //we dont need to do this
    }
    
    res.setHeader('Content-Type', 'text/csv')
    res.setHeader('Content-Disposition', 'attachment; filename="output.csv"');
    workbook.csv.write(res);
    // Clean up uploaded file
   
   
  } catch (error) {
    err('Error processing file:' + error);
    res.status(500).send(error);
    // Ensure uploaded file is cleaned up in case of an error
    if (req.file && req.file.path) {
      fs.unlink(req.file.path, (fr) => {
        if (fr) {
          err('Error deleting file:' + fr);
        }
      });
    }
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
