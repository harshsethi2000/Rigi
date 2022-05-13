const express = require("express");
const req = require("express/lib/request");
//googleapis
const { google } = require("googleapis");
//initilize express
const app = express();
const PORT=3000;

//set app view engine
app.set('view engine','ejs')
app.set('views','./views')
const reader =require('xlsx');

//function to download the excel file send on email by the user and to download that file in our working folder
downloadAndUpdate();
//function to read from sheet and transform two tabs to one array based on id
var finalArray=readAndTransform();
//function to authenticate then update the google sheet
authenticateAndUpdateSheet(finalArray);



function downloadAndUpdate(){

var fs = require("fs");
var buffer = require("buffer");
var Imap = require("imap");
const base64 = require('base64-stream')

var imap = new Imap({
  user: "harhsethi2000@gmail.com",
  password: "*****",
  host: "imap.gmail.com",
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false }

 
});

function toUpper(thing) {
  return thing && thing.toUpperCase ? thing.toUpperCase() : thing;
}

function findAttachmentParts(struct, attachments) {
  attachments = attachments || [];
  for (var i = 0, len = struct.length, r; i < len; ++i) {
    if (Array.isArray(struct[i])) {
      findAttachmentParts(struct[i], attachments);
    } else {
      if (
        struct[i].disposition &&
        ["INLINE", "ATTACHMENT"].indexOf(toUpper(struct[i].disposition.type)) >
          -1
      ) {
        attachments.push(struct[i]);
      }
    }
  }
  return attachments;
}

function buildAttMessageFunction(attachment) {
  var filename = attachment.params.name;
  var encoding = attachment.encoding;
  

  return function(msg, seqno) {
    var prefix = "(#" + seqno + ") ";
    msg.on("body", function(stream, info) {
      ;
      //Create a write stream so that we can stream the attachment to file;
     
      var writeStream = fs.createWriteStream('2'+filename);
      writeStream.on("finish", function() {
        
      });

      // stream.pipe(writeStream); this would write base64 data to the file.
      // so we decode during streaming using
      if (toUpper(encoding) === "BASE64") {
        
        if (encoding === 'BASE64') stream.pipe(new base64.Base64Decode()).pipe(writeStream)
      }
    });
    msg.once("end", function() {
      console.log(prefix + "Finished attachment %s", filename);
    });
  };
}

imap.once("ready", function() {
  imap.openBox("INBOX", true, function(err, box) {
    if (err) throw err;
    var f = imap.seq.fetch("1:10", {
      bodies: ["HEADER.FIELDS (FROM TO SUBJECT DATE)"],
      struct: true
    });
    f.on("message", function(msg, seqno) {
      
      var prefix = "(#" + seqno + ") ";
      msg.on("body", function(stream, info) {
        var buffer = "";
        stream.on("data", function(chunk) {
          buffer += chunk.toString("utf8");
        });
        stream.once("end", function() {
          console.log(prefix + "Parsed header: %s", Imap.parseHeader(buffer));
        });
      });
      msg.once("attributes", function(attrs) {
        var attachments = findAttachmentParts(attrs.struct);
        console.log(prefix + "Has attachments: %d", attachments.length);
        for (var i = 0, len = attachments.length; i < len; ++i) {
          var attachment = attachments[i];
          console.log("Attachment name is "+attachment.params.name);
          if(attachment.params.name.includes(".xlsx")==false)continue;
          
          var f = imap.fetch(attrs.uid, {
            //do not use imap.seq.fetch here
            bodies: [attachment.partID],
            struct: true
          });
          //build function to process attachment message
          f.on("message", buildAttMessageFunction(attachment));
        }
      });
      msg.once("end", function() {
        //console.log(prefix + "Finished email");
      });
    });
    f.once("error", function(err) {
      //console.log("Fetch error: " + err);
    });
    f.once("end", function() {
      //console.log("Done fetching all attachments ");
      imap.end();
    });
  });
});

imap.once("error", function(err) {
  console.log(err);
});

imap.once("end", function() {
  console.log("Connection ended");
});

imap.connect();

}


function readAndTransform(){

// const fileName="Input report-2022-05-12.xlsx";
// const file = reader.readFile('./Input report-2022-05-12.xlsx')

const fileName="Input report-2021-07-21.xlsx";
const file = reader.readFile('./Input report-2021-07-21.xlsx')

//get date from fileName
var date="";
for(var i=13;i<23;i++)
{
date=date+fileName[i];
}


let tmpData = []
let tmpDataTwo=[];
  
const sheets = file.SheetNames
  
//iterate over the sheets and store the data in our temporary variables
for(let i = 0; i < sheets.length; i++)
{
   const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
       if(i==1)tmpData.push(res);
       else tmpDataTwo.push(res);
   })
}



//function to merge two arrays based on id
const mergeArrays = (arr1 = [], arr2 = []) => {
    let res = [];
    res = arr1.map(obj => {
       const index = arr2.findIndex(el => el["id"] == obj["id"]);
       const {amount} = index !== -1 ? arr2[index] : {};
       return {
          ...obj,
          amount
       };
    });
    return res;
 };

 //merge two array based on id
var finalArray=mergeArrays(tmpData,tmpDataTwo);
//finalArray will contain all the data that we have to add in our google sheets


//add date field to an array
finalArray.forEach((elem)=>{
    elem["date"]=date;
    });
//add breakage point for this file    
    var tmpJson={
        id:'end',
        user_name:'end',
        amount:'end',
        date:'end'
    }
finalArray[finalArray.length]=tmpJson;

//console.log(finalArray);
return finalArray;
}

//authenticate the google api

function authenticateAndUpdateSheet(finalArray)
{

const auth = new google.auth.GoogleAuth({
    keyFile: "hs.json", //the key file
    //url to spreadsheets API
    scopes: "https://www.googleapis.com/auth/spreadsheets", 
});



app.get('/appendDataToSheet',async(req,res)=>{
    //call the helper function
    helper();
    res.render('addedSuccessfully')
});

//helper function to 
async function helper() {
    
const authClientObject = await auth.getClient();
const googleSheetsInstance = google.sheets({ version: "v4", auth: authClientObject });
//mention the googleSheetId
const spreadsheetId = "1QYEHGQjTlAxTICmnK8NI7gBmVT1axz_WI_p7tA6zDvg";

//iterate over array and append each values to our google sheets
for(var i=0;i<finalArray.length;i++)
{



await googleSheetsInstance.spreadsheets.values.append({
        auth, //auth object
        spreadsheetId, //spreadsheet id
        range: "Sheet1!A:B", //sheet name and range of cells
        valueInputOption: "USER_ENTERED", // The information will be passed according to what the usere passes in as date, number or text
        resource: {

            values: [[finalArray[i].id,finalArray[i].user_name,finalArray[i].amount,finalArray[i].date]],
        },
    });

}
    
}


}

app.listen(PORT,()=>console.log(`Server running on Port: http://localhost:${PORT}`));





