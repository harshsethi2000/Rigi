//include all the necessary package
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

// const fileName="Input report-2021-07-21.xlsx";
// const file = reader.readFile('./Input report-2021-07-21.xlsx')


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


//authenticate the google api

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


app.listen(PORT,()=>console.log(`Server running on Port: http://localhost:${PORT}`));





// //requiring path and fs modules
// const path = require('path');
// const fs = require('fs');
// //joining path of directory 
// const directoryPath = path.join(__dirname, '');
// //passsing directoryPath and callback function
// fs.readdir(directoryPath, function (err, files) {
//     //handling error
//     if (err) {
//         return console.log('Unable to scan directory: ' + err);
//     } 
//     //listing all files using forEach
//     files.forEach(function (file) {
//         // Do whatever you want to do with the file
//         console.log(file); 
//     });
// });





//  // import reader from 'xlsx'
// // Reading our test file
// //import fs, { appendFile } from 'fs';

// // import google from 'googleapis';
// // import express from 'express';
// const express=require('express');
// const {google}=require('googleapis');



// const authentication = async()=>{
//     const auth =new google.auth.GoogleAuth({
//         keyFile:"hs.json",
//         scopes:"https://www.googleapis.com/auth/spreadsheets"
//     });

// const client=await auth.getClient();
// const s=google.sheets({
// version:'v4',
// auth: client
// });
// return {s};
    
// }




// const id="1QYEHGQjTlAxTICmnK8NI7gBmVT1axz_WI_p7tA6zDvg";
// app.get('/',async (req,res)=>{
//     try
//     {
//         const {sheets}=await authentication();
//         const response=await sheets.spreadsheetId.value.get({
//             spreadsheetId:id,
//             range:'Sheet1,'

//         })
        
//         res.send(response.data)

//     }catch(error)
//     {
//         console.log(error);
//     }
    

//     });

    
// app.listen(PORT,()=>console.log(`Server running on Port: http://localhost:${PORT}`));




// // const fileName="Input report-2021-07-21.xlsx";
// // var date="";
// // for(var i=13;i<23;i++)
// // {
// // date=date+fileName[i];
// // }

// // const file = reader.readFile('./Input report-2021-07-21.xlsx')
  
// // let data = []
// // let data2=[];
  
// // const sheets = file.SheetNames
  

// // //    const temp = reader.utils.sheet_to_json(
// // //         file.Sheets[file.SheetNames[1]])
// // //    temp.forEach((res) => {
// // //       data.push(res)
// // //    })

// // //      const tempTwo = reader.utils.sheet_to_json(
// // //         file.Sheets[file.SheetNames[2]])
// // //    tempTwo.forEach((res) => {
// // //       data2.push(res)
// // //    })

// // for(let i = 0; i < sheets.length; i++)
// // {
// //    const temp = reader.utils.sheet_to_json(
// //         file.Sheets[file.SheetNames[i]])
// //    temp.forEach((res) => {
// //        if(i==1)data.push(res);
// //        else data2.push(res);
// //    })
// // }





// // //merge two array based on id

// // const mergeArrays = (arr1 = [], arr2 = []) => {
// //     let res = [];
// //     res = arr1.map(obj => {
// //        const index = arr2.findIndex(el => el["id"] == obj["id"]);
// //        const {amount} = index !== -1 ? arr2[index] : {};
// //        return {
// //           ...obj,
// //           amount
// //        };
// //     });
// //     return res;
// //  };

// // var finalArray=mergeArrays(data,data2);

// // //add date field to an array
// // finalArray.forEach((elem)=>{
// //     elem["date"]=date;
// //     });

// //     var tmpJson={
// //         "id": '',
// //         user_name: '',
// //         amount: '',
// //         date: '',
// //     }
// // finalArray[finalArray.length]=tmpJson;

// // const ws = reader.utils.json_to_sheet(finalArray)
  
// // reader.utils.book_append_sheet(file,ws,"Sheet3")
  
// // // Writing to our file
// // reader.writeFile(file,'./testThree.xlsx')

// // // fs.appendFileSync('./testThree.xlsx', finalData);
  
// // // Printing data
// // console.log(data)
// // console.log(data2)
// // console.log(finalArray);
// // console.log(date);