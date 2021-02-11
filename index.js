const express=require("express");
let app=express();
let port=3000;
var cors = require('cors');
//require('dotenv').config();

app.use(express.json());
app.use(cors());

let XLSX=require("xlsx");
let workbook = XLSX.readFile("stocks.xlsx");
let sheet_name_list = workbook.SheetNames;
//console.log(workbook);
//console.log(sheet_name_list); // getting as Sheet1


app.get('/',(req,res)=>{
  let y="Sheet1"
  //console.log(y);  
  var worksheet = workbook.Sheets[y];
  //getting the complete sheet
  //console.log(worksheet);
  
  var headers = {};
  var data = [];
  //res.send(worksheet);
  for (z in worksheet) {
    //console.log(z);
    if (z[0] === "!") continue;
    //parse out the column, row, and value
    var col = z.substring(0, 1);
    // console.log(col);

    var row = parseInt(z.substring(1));
    // console.log(row);

    var value = worksheet[z].v;
    // console.log(value);

    //store header names
    if (row == 1) {
      headers[col] = value;
      // storing the header names
      continue;
    }

    if (!data[row]) data[row] = {};
    data[row][headers[col]] = value;
  }
  //drop those first two rows which are empty
  data.shift();
  data.shift();
  //data.shift();
  res.send(data);
  //return;
    //res.send(workbook.Sheets["Sheet1"]);
});

app.listen(port,()=>{
  console.log(`example app listening at https://localhost:${port}`);
});