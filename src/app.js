const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const ExcelJS = require('exceljs');
const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

const fastcsv = require('fast-csv');

const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;


const Stream = require('stream'); //this language is so fucked


const IS_LOCAL = true;

/*
need to check size client side
need to add loading client side
need to make preview good and optional


should be able to convert file w/out running out of memory! 

*/

// Middleware to handle form data

const max_upload_size_mb = 40;

async function big_excel_to_csv(file){

  const workbook = new ExcelJS.Workbook();
  const csvWriter = fastcsv.format({ headers: true }); 
  const passThroughStream = new Stream.PassThrough();
  csvWriter.pipe(passThroughStream);
  await workbook.xlsx.read(file).then((wb) => {
    log("loaded excel!");
    const sheet = wb.getWorksheet();
    sheet.eachRow((row) => {
      csvWriter.write(row.values);
    });


    csvWriter.end();
  });
  return passThroughStream;
}




async function big_excel_to_csv2(file){

  const workbook = new ExcelJS.Workbook();
  const csvWriter = fastcsv.format({ headers: true }); 
  const passThroughStream = new Stream.PassThrough();
  csvWriter.pipe(passThroughStream);


  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(file);
  let c = 0;
  let r = 0;
  for await (const worksheetReader of workbookReader) { //fucked if more than 1 worksheet
    log("reading worksheet #" + c.toString());
    c += 1;
    for await (const row of worksheetReader) {
      const { values } = row;
      values.shift();
      
      csvWriter.write(values);    
      r += 1;
    }
    console.log(`wrote ${r} rows`);
    csvWriter.end(); //should just break here so we never crash on >1 worksheet

    //break; //^
  }
  return passThroughStream;
}



const upload = multer({ 
  dest: 'uploads/',
limits: { fileSize: max_upload_size_mb * 1024 * 1024 } 
 });

 function getCellResult(worksheet, cellLabel) { //formula edge cases ! 
  if (worksheet.getCell(cellLabel).formula) {
    return parser.parse(worksheet.getCell(cellLabel).formula).result;
  } else {
    return worksheet.getCell(cellCoord.label).value;
  }
}
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


if(IS_LOCAL)
{
    // Serve static files
  app.use(express.static(path.join(__dirname, '../public')));
  app.use(bodyParser.urlencoded({ extended: true }));
}



app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
    res.setHeader('Content-Type', 'text/plain');
    return res.status(413).send('File size exceeds the limit of ' + max_upload_size_mb.toString() + ' MB!');
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
 
   /*
    var options = { filename: req.file.path, useStyles: true, useSharedStrings: true }; 
    var workbook = new ExcelJS.stream.xlsx.WorkbookReader(req.file.path);
    for await (const worksheetReader of workbook) {
      console.log("piss");
      res.setHeader('Content-Type', 'text/csv')
      res.setHeader('Content-Disposition', 'attachment; filename="output.csv"');
      
    
     
      await worksheetReader.csv.write(res, {stream: true});

      return;}
    }*/

const size_threshold = 1;


  if(req.file.size > size_threshold ){
    log("using large file method");


    res.setHeader('Content-Type', 'text/csv')
    res.setHeader('Content-Disposition', 'attachment; filename="output.csv"');
    const fd = fs.createReadStream(req.file.path);
    const c = await big_excel_to_csv2(fd); 

    log("sending response!");
    //we should stream this back somehow but instead we do this shit and it sucks a bit 


    const chunks = [];

    c.on('readable', () => {
      let chunk;
      while (null !== (chunk = c.read())) {
        chunks.push(chunk);
      }
    });

    c.on('end', () => {
      const content = chunks.join('');
      res.send(content);
    });
  
    //res.send(content);

  } else{ //quick n dirty
    log("using quick n dirty method");


    const workbook = new ExcelJS.Workbook();
    await workbook.readFile(req.file.stream);
    const worksheet = workbook.getWorksheet(0);
    if (!worksheet){
      err("null worksheet");
      throw new err('Worksheet not found in uploaded workbook');
    } 

    log(`A1 = ${worksheet.getCell("A1").value}`);

    res.setHeader('Content-Type', 'text/csv')
    res.setHeader('Content-Disposition', 'attachment; filename="output.csv"');
    await workbook.csv.write(res, {stream: true});
  }
    
    //await workbook.xlsx.readFile(req.file.path);
    //i think we need to do a streaming thing! 


    //https://gist.github.com/davidhq/0afed70985842cac6fc4e00f88a71bd2
   

    
  
  
  // Clean up uploaded file
  log("deleting uploaded file " + req.file.path);
  fs.unlink(req.file.path, (error) => {
      if (error) {
        err('Error unlinking file: ' + error);
      } else {
        log('upload returned and deleted successfully. \n');
      }
    });
   
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
