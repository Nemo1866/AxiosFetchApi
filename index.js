const { default: axios } = require("axios");
const express = require("express");
const app = express();
const fs = require("fs");
const path = require("path");
const csvtojson = require("csvtojson");
const csv = require("csv-parser");

const ExcelJS = require("exceljs");



const params2 = {
  user: "mweiran",
  password: "Sectran13",
  mode: "login",
};
app.get("/", async (req, res) => {
  try {
    const LoginData = await axios.post(
      "https://sectranws.apunix.com/summit/SrvManager",
      params2,
      {
        withCredentials: true,
        headers: {
          "Content-Type": "multipart/form-data",
        },
      }
    );

    const kuId = LoginData.headers["set-cookie"][2].split(";")[0];
    const token = LoginData.headers["set-cookie"][3].split(";")[0];
    const cookieData = `PrefRegion=America; SelectedRegion=America; SelectedCustomer=46; SelectedLang=null; kcuname=mweiran; ${kuId}; ${token}`;
    const customers = [46, 59, 60, 61, 62];
    const responses = await Promise.all(
      customers.map(async (customer) => {
        const response = await axios.post(
          "https://sectranws.apunix.com/summit/SrvRunReports",
          {
            site: 0,
            city: "",
            region: "America",
            reportType: "",
            eod_start_date: "2023-03-25",
            eod_date: "2023-03-25",
            armored_start_date: "2023-03-25",
            armored_date: "2023-03-25",
            custom_type: "001%",
            custom_date1: "2023-03-25",
            custom_date2: "2023-03-25",
            provisional_history: 1,
            provisional_date: "2023-03-25",
            provisional_time: "16:00:00",
            user_type: "%",
            user_date1: "2023-03-25",
            user_date2: "2023-03-25",
            error_date1: "2023-03-25",
            error_date2: "2023-03-25",
            siteuser_date1: "2023-03-25",
            siteuser_date2: "2023-03-25",
            con_cust: customer,
            reportType2: "con_eod",
            con_cr_start: "2023-03-25",
            con_cr_end: "2023-03-25",
            con_eod_start: "2023-01-24",
            con_eod_end: "2023-01-25",
            mode: "report",
            user: "mweiran",
            access_level: 4,
            accessor: "Maria Weiran (FFCorp)",
            ReportDate: "2023-01-24",
            ReportTime: "",
            ReportDate2: "2023-01-25",
            ReportCity: "",
            ReportSite: "",
            ReportHost: "",
            ReportName: "con_eod",
            ReportFilter1: customer,
            ReportFilter2: "00:00",
            ReportRegion: "America",
            ReportDays: "",
            report_file_type: "CSV",
          },
          {
            responseType: "stream",
            withCredentials: true,
            headers: {
              "Content-Type": "multipart/form-data",
              Cookie: cookieData,
            },
          }
        );

        const contentDispositionHeader =
          response.headers["content-disposition"];
        const fileName = contentDispositionHeader
          ? contentDispositionHeader.split("=")[1]
          : "couldNotDownlad.csv";
        const filePath = path.join(__dirname, fileName);

        const stream = await response.data.pipe(fs.createWriteStream(filePath));
        // stream.on("finish", () => {
        //   console.log(`File ${fileName} has been downloaded successfully.`)
        //   return res.json({message:"Sucessfully Downloaded"})
        // })

        return {
          message: `File ${fileName} has been downloaded successfully.`,
        };
      })
    );
      //  const jsonArray=[]
    const test = responses.map((message) => message.message.split(" ")[1]);
    // for (let message of test) {
    //   if (message) {
    //     const jsonFilePath = `./${message.replace('.csv', `.json`)}`;
    //     await convertCsvToJson(message, jsonFilePath);
    //     jsonArray.push(jsonFilePath);
    //   }
    // }
    // await convertJsonsToExcel(jsonArray,"./Template.xlsx")
    // await convertCsvToJson(test)
    await processCsvFiles(test, "./Template.xlsx");

    res.json({ Added: "Added Sucessfully" });
  } catch (error) {
    console.log(error);
    res.status(500).send("Error downloading file");
  }
});

// async function convertCsvToJson(csvFilePath, jsonFilePath) {
//   return new Promise((resolve, reject) => {
//     const jsonArray = [];
//     fs.createReadStream(csvFilePath)
//       .pipe(csv())
//       .on('data', (data) => {
//         jsonArray.push(data);
//       })
//       .on('end', () => {
//         fs.writeFile(jsonFilePath, JSON.stringify(jsonArray), (err) => {
//           if (err) {
//             console.error('Error writing JSON file:', err);
//             reject(err);
//           } else {
//             console.log(`CSV file ${csvFilePath} converted to JSON file ${jsonFilePath}`);
//             resolve();
//           }
//         });
//       })
//       .on('error', (err) => {
//         console.error('Error reading CSV file:', err);
//         reject(err);
//       });
//   });
// }



// async function convertJsonsToExcel(jsonFilePaths, excelFilePath) {
//   const workbook = new ExcelJS.Workbook();

//   // Read the existing workbook from the file
//   await workbook.xlsx.readFile(excelFilePath);

//   // Get the worksheet named "Cash" or create a new one if it doesn't exist
//   let worksheet = workbook.getWorksheet('Cash');
//   if (!worksheet) {
//     worksheet = workbook.addWorksheet('Cash');
//   }

//   // Get the last row index in the worksheet
//   const lastRowIndex = worksheet.lastRow ? worksheet.lastRow.number : 0;
//   let isFirstFile=true
//   // Add the JSON data to the worksheet
//   for (let i = 0; i < jsonFilePaths.length; i++) {
//     const jsonFilePath = jsonFilePaths[i];

//     const jsonArray = await fs.promises.readFile(jsonFilePath).then(data => JSON.parse(data));

//     // if (i===0) {

//     // // Add header row from the first file
//     // const headers = Object.values(jsonArray[0]);

//     // worksheet.addRow(headers);

//     // }
//     // Add data rows
//     jsonArray.forEach((row, index) => {
//       if(i===0 || index!==0){
//       const values = Object.values(row);

//       worksheet.addRow(values, lastRowIndex + index + 1);
//       }
//     });

//     console.log(`JSON file ${jsonFilePath} appended to worksheet 'Cash'`);
//   }

//   // Save the updated workbook to the file
//   await workbook.xlsx.writeFile(excelFilePath);
//   console.log(`JSON files appended to Excel file ${excelFilePath}`);
// }


// async function processCsvFiles(csvFiles = [], excelFilePath) {
//   try {
//     const workbook = new ExcelJS.Workbook();
//     await workbook.xlsx.readFile(excelFilePath);
//     const worksheet = workbook.getWorksheet('Cash');
//     let currentRow = 1;
//     let headerAdded = false; // move outside of the loop
//     })
    // for (const [index, csvFile] of csvFiles.entries()) {
    //   let currentLine = 1;
    //   await new Promise((resolve, reject) => {
    //     fs.createReadStream(csvFile)
    //       .pipe(csv())
    //       .on('data', (data) => {
    //         console.log(data);
    //         if (index === 0 && currentLine <= 1) { 
    //           if (!headerAdded) {
    //             // Add header row to worksheet
    //             // Object.keys(data).forEach((key, index) => {
    //             //   worksheet.getCell(currentRow, index + 3).value = key;
    //             // });
    //             // currentRow++;
    //             // Add data to the worksheet

    //           Object.values(data).forEach((value, index) => {
    //             const test=worksheet.getCell(currentRow, index + 3).value = value;
  
    //           });
    //           // console.log(rearranged);
           
    //           currentRow++;
    //             headerAdded = true;
    //           }
    //           // // Add data to the worksheet
    //           // Object.values(data).forEach((value, index) => {
    //           //   worksheet.getCell(currentRow, index + 3).value = value;
    //           // });
    //           // currentRow++;
    //         } else if (currentLine > 2) { // for all other csv files, skip the headers
    //           // Add data to the worksheet
    //           Object.values(data).forEach((value, index) => {
    //             worksheet.getCell(currentRow, index + 3).value = value;
    //           });
    //           currentRow++;
    //         }
    //         currentLine++;
    //       })
    //       .on('end', () => {
    //         console.log(`CSV file ${csvFile} processed successfully`);
    //         resolve();
    //       })
    //       .on('error', (err) => {
    //         console.error(`Error processing CSV file ${csvFile}:`, err);
    //         reject(err);
    //       });
    //   });
    // }
//     await workbook.xlsx.writeFile(excelFilePath);
//     console.log(`Data added to Excel file ${excelFilePath}`);
//   } catch (error) {
//     console.error('Error processing CSV files:', error);
//     throw error;
//   }
// }




// async function processCsvFiles(csvFiles = [], excelFilePath) {
//   try {
//     const workbook = new ExcelJS.Workbook();

//     await workbook.xlsx.readFile(excelFilePath);
//     let worksheet = workbook.getWorksheet('Cash2');
//     if(!worksheet){
//       worksheet = workbook.addWorksheet("Cash2")
//     }
   


  
//     let headerAdded = false;

//     for (const file of csvFiles) {
//       const content = await fs.promises.readFile(file, 'utf-8');
//       const lines = content.split('\n');

//       if (!headerAdded) {
//         let headers = lines[1].split(',').map(header=>header.replace(/"/g,''));
//         let row=worksheet.addRow()
       
//       for(let j=0;j<headers.length;j++){
//         row.getCell(j+3).value=headers[j]
//       }
//         headerAdded = true;
//       }

//       for (let i = 2; i < lines.length; i++) {
//         let values = lines[i].split(',').map(value=>value.replace(/"/g,''))
//         let row = worksheet.addRow();
  
//         for (let j = 0; j < values.length; j++) { // start from second column
//           row.getCell(j+3).value = values[j]; // shift to third column
//         }
//       }
//       console.log(`CSV File ${file} have been processed successfully`);
//     }      

//     await workbook.xlsx.writeFile(excelFilePath);
//     console.log('All CSV files have been processed successfully.');
//   } catch (error) {
//     console.error('Error occurred while processing CSV files:', error);
//   }
// }




async function processCsvFiles(csvFiles = [], excelFilePath) {
  try {
    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(excelFilePath);
    let worksheet = workbook.getWorksheet('Cash2');
    if(!worksheet){
      worksheet = workbook.addWorksheet("Cash2")
    }

    let cashWorksheet = workbook.getWorksheet('Cash');
    if(!cashWorksheet){
      console.log('Cash worksheet not found');
      return;
    }

    let headerAdded = false;
 

    for (const file of csvFiles) {
      const content = await fs.promises.readFile(file, 'utf-8');
      const lines = content.split('\n');

      if (!headerAdded) {
        let headers = lines[1].split(',').map(header=>header.replace(/"/g,''));
        let row=worksheet.addRow();

      
        row.getCell(1).value = cashWorksheet.getCell(1,1).value
        row.getCell(1).formula = cashWorksheet.getCell(1,1).formula
        row.getCell(2).value = cashWorksheet.getCell(1,2).value;
        row.getCell(2).formula = cashWorksheet.getCell(1,2).formula;
        row.getCell(2).style = cashWorksheet.getCell(1,2).style;
        row.getCell(34).value=cashWorksheet.getCell(1,34).value;
        row.getCell(34).formula=cashWorksheet.getCell(1,34).formula;
        row.getCell(34).style=cashWorksheet.getCell(1,34).style;

 
        
        for(let j=0;j<headers.length;j++){
          row.getCell(j+3).value=headers[j]
        }
        headerAdded = true;
      }



      for (let i = 2; i < lines.length; i++) {
        let values = lines[i].split(',').map(value=>value.replace(/"/g,''))
        let row = worksheet.addRow();

       
        row.getCell(1).value = cashWorksheet.getCell(i,1).value;
        row.getCell(1).formula = cashWorksheet.getCell(i,1).formula;
        row.getCell(2).value = cashWorksheet.getCell(i,2).value;
        row.getCell(2).formula = cashWorksheet.getCell(i,2).formula;
        row.getCell(2).style = cashWorksheet.getCell(i,2).style; 
        row.getCell(34).value=cashWorksheet.getCell(i,34).value  
        row.getCell(34).formula=cashWorksheet.getCell(i,34).formula  
        row.getCell(34).style=cashWorksheet.getCell(i,34).style  

        for (let j = 0; j < values.length; j++) { 
          row.getCell(j+3).value = values[j]; 
        }
      }
     
      console.log(`CSV File ${file} has been processed successfully`);
   
    }      
   workbook.calcProperties.fullCalcOnLoad=true
   if(cashWorksheet){
   workbook.removeWorksheet("Cash")
   worksheet.name="Cash"
   }
    await workbook.xlsx.writeFile(excelFilePath);
    console.log('All CSV files have been processed successfully.');
  } catch (error) {
    console.error('Error occurred while processing CSV files:', error);
  }
}



app.listen(3000, () => console.log("Server is running on port 3000"));
