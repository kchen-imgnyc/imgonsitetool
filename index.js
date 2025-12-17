let outstandingData=new Map();
let allData=[];
let locationChoice="NYC";

document.addEventListener("DOMContentLoaded", function() {
 
const excelInputBtn=document.getElementById("excel-file-input")
const fileNameDisplay = document.getElementById('file-name-display');

excelInputBtn.addEventListener(("change"),function(){
    const files = event.target.files;

    if (files.length > 0) {
        allData=[]
        const selectedFile = files[0];
        fileNameDisplay.textContent = selectedFile.name;
        console.log('File Name:', selectedFile.name);

        const reader = new FileReader();
        reader.onload = function(e) {
        const data = e.target.result;

        const workbook = XLSX.read(data, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        document.getElementById('output').textContent = JSON.stringify(json, null, 2);
        
        addedCodes=[]

        for(let i=0;i<json.length;i++){
            item=json[i]
                
            icode=item["ICode"]
            discription=item["Description"]
            unitCost=item["UnitCost"]
            manufacturer=item["Manufacturer"]

            if (addedCodes.indexOf(icode) == -1 && icode != undefined){
                allData.push([icode,discription,unitCost,manufacturer])
                addedCodes.push(icode)
            }
            else if( icode == undefined){
                allData.push([icode,discription,unitCost,manufacturer])
                addedCodes.push(icode)
            }
            
            

        }

        console.log(allData)

        };

        reader.readAsArrayBuffer(selectedFile);



    } else {

        fileNameDisplay.textContent = 'No file selected';
    }
         
});



const outstandingExcelBtn=document.getElementById("excel-file-input-outstanding")
const fileNameDisplayAsset=document.getElementById('file-name-display-asset');

outstandingExcelBtn.addEventListener(("change"),function(){

    const files = event.target.files;
    if (files.length > 0) {
        outstandingData=new Map()
        const selectedFile = files[0];
        fileNameDisplayAsset.textContent = selectedFile.name;
    
        const reader = new FileReader();
        reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);


        for(let i=0;i<json.length;i++){
            item=json[i]

            icode=item["ICode"]
            category=item["Category"]
            imageID=item["ImageId"]
            qty=item["Quantity"]

            dataObj={
                "icode":icode,
                "category":category,
                "imageID":imageID,
                "qty":qty
            }
            
            if (outstandingData.has(icode)){
                value=outstandingData.get(icode)
                value["qty"]=value["qty"] +1
                outstandingData.set(icode,value)
            }
            else{
                outstandingData.set(icode,dataObj)
            }

        }

        console.log(outstandingData)


        };

        reader.readAsArrayBuffer(selectedFile);

    } else {

        fileNameDisplayAsset.textContent = 'No file selected';
    }
         

});


let icodesRan=[]

const generateReportButton =document.getElementById("process_files")

generateReportButton.addEventListener("click", async function() {
    // 1. Initial Checks
    icodesRan=[]
    if (allData.length == 0) {
        return alert("Please Upload All Items Excel Sheet");
    }

    if (outstandingData.size == 0) {
        return alert("Please Upload Outstanding Excel");
    }

    const filename = "onsite.xlsx";

    // 2. Setup Workbook and Worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    // Define Header (used for columns and the first row)
    const header = [
        "ICODE", "QTY", "ROOM + ITEM", "ESTIMATE", "ESTIMATE TOTAL", "COST", 
        "TOTAL", "PRICE", "VENDOR + ITEM NAME", "ITEM DESCRIPTION", "IMAGE", 
        "REMOVAL NOTES"
    ];

    // Define columns for width and formatting (optional but good practice)
    worksheet.columns = [
        { 
            header: header[0], 
            key: 'icode', 
            width: 10,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[1], 
            key: 'qty', 
            width: 8,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[2], 
            key: 'room', 
            width: 25,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[3], 
            key: 'estimate', 
            width: 15, 
            style: { 
                numFmt: '$#,##0.00', 
                alignment: { horizontal: 'center' }, 
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[4], 
            key: 'est_total', 
            width: 18, 
            style: { 
                numFmt: '$#,##0.00',
                alignment: { horizontal: 'center' }, 
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[5], 
            key: 'cost', 
            width: 15, 
            style: { 
                numFmt: '$#,##0.00',
                alignment: { horizontal: 'center' }, 
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[6], 
            key: 'total', 
            width: 18, 
            style: { 
                numFmt: '$#,##0.00',
                alignment: { horizontal: 'center' }, 
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[7], 
            key: 'price', 
            width: 15, 
            style: { 
                numFmt: '$#,##0.00',
                alignment: { horizontal: 'center' }, 
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[8], 
            key: 'vendor_item', 
            width: 40,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[9], 
            key: 'description', 
            width: 40,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[10], 
            key: 'image', 
            width: 20,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        },
        { 
            header: header[11], 
            key: 'notes', 
            width: 25,
            style: { 
                alignment: { horizontal: 'center' },
                font: { size: 10, name: 'Montserrat' } // Font updated
            } 
        }
    ];
    


    let row_num = 2; 
    for (let i = 0; i < allData.length; i++) {

        let row = allData[i];
        let icode = row[0];
        let description = row[1];
        let unitCost = row[2];
        let manufacturer = row[3];

        let extraData = outstandingData.get(icode);
        let category = "";
        let imageID = "";
        let qty = 0;

        if (icode != undefined && extraData == undefined){
            continue
        }

        if(icode !== undefined && icodesRan.indexOf(icode) !== -1){
            continue
        }


        let estimateFormula = `B${row_num}*D${row_num}`;
        let totalFormula = `B${row_num}*F${row_num}`;

        let imageLink = '';

        if (manufacturer != undefined){
            if (manufacturer == "IMG") {
                manufacturer = "IMG ART LOFT:";
            } else if (manufacturer == "CUSTOM IMG") {
                manufacturer = "IMG CUSTOM";
            } else if ( manufacturer.length > 0 && manufacturer.length <= 3) {
                manufacturer = "IMG HOME EXCLUSIVE:";
            }
        }

        if (extraData != undefined) {
            category = extraData["category"];
            imageID = extraData["imageID"];
            qty = extraData["qty"];
        }
        
        if (imageID != undefined) {
            imageLink = `IMAGE("https://imgnyc.rentalworks.cloud/api/v1/appimage/getimage?appimageid=${String(imageID)}&thumbnail=false",4,150,150)`;
        }

        // The array format for adding a row
        let newRowValues = [
            icode, qty, category, unitCost, estimateFormula, unitCost, 
            totalFormula, unitCost, `${manufacturer} - ${description}`, "", // description is column J (index 9)
            imageLink, "" // image is column K (index 10), removal notes is L (index 11)
        ];


        if (icode == undefined) {
            worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 1
            worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 2
            newRowValues = [
                "", "", description, "", "", "", 
                "", "", "", "", 
                "", ""
            ];
            row_num = row_num + 2
        }


        // Handle missing ImageID
        if (imageID == undefined || imageID == "") {
            newRowValues[10] = "";
        }
        
        // Add the row
        let newRow = worksheet.addRow(newRowValues);
        icodesRan.push(icode)
        row_num = row_num + 1

        // Set formulas in ExcelJS *after* adding the row, using the cell object
        if (icode != undefined) {
            // Apply formula to ESTIMATE TOTAL (Column E, index 4)
            newRow.getCell(5).value = { formula: estimateFormula }; 
            
            // Apply formula to TOTAL (Column G, index 6)
            newRow.getCell(7).value = { formula: totalFormula }; 
            
            // Apply image formula to IMAGE (Column K, index 10)
            if (imageID != undefined) {
               newRow.getCell(11).value = { formula: imageLink };
               newRow.height = 112.5; // 150px is approx 112.5 points in Excel
            }
        }
    }

    worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 1
    worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 2
    row_num = row_num + 2    

    //`B${row_num}*D${row_num}`

    if (locationChoice === "NYC"){
        ThirdLastRow=[
                    "",
                    "",
                    "", 
                    "SubTotal:", 
                    `SUM(E1:E${row_num-1})`, 
                    "", 
                    `SUM(G1:G${row_num-1})`, 
                    `SUM(H1:H${row_num-1})`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];

        newRow = worksheet.addRow(ThirdLastRow);
        newRow.getCell(5).value = { formula: `SUM(E1:E${row_num-1})` }
        newRow.getCell(7).value ={ formula: `SUM(G1:G${row_num-1})` }
        newRow.getCell(8).value ={ formula: `SUM(H1:H${row_num-1})` }

        row_num =row_num +1

        SecondLastRow=[
                    "",
                    "",
                    "", 
                    "NY Sales Tax:", 
                    `E${row_num-1}*0.008875`, 
                    "", 
                    `G${row_num-1}*0.008875`, 
                    `H${row_num-1}*0.008875`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];

        newRow = worksheet.addRow(SecondLastRow);
        newRow.getCell(5).value = { formula: `E${row_num-1}*0.008875` }
        newRow.getCell(7).value ={ formula: `G${row_num-1}*0.008875` }
        newRow.getCell(8).value ={ formula: `H${row_num-1}*0.008875` }
        row_num =row_num +1

        LastRow=[
                    "",
                    "",
                    "",
                    "Total:", 
                    `E${row_num-2}+E${row_num-1}`, 
                    "",
                    `G${row_num-2}+G${row_num-1}`, 
                    `H${row_num-2}+H${row_num-1}`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];
        newRow = worksheet.addRow(LastRow);
        newRow.getCell(5).value = { formula: `E${row_num-2}+E${row_num-1}` }
        newRow.getCell(7).value ={ formula: `G${row_num-2}+G${row_num-1}` }
        newRow.getCell(8).value ={ formula: `H${row_num-2}+H${row_num-1}` }


    
    }
    else if(locationChoice === "FL"){
         ThirdLastRow=[
                    "",
                    "",
                    "", 
                    "SubTotal:", 
                    `SUM(E1:E${row_num-1})`, 
                    "", 
                    `SUM(G1:G${row_num-1})`, 
                    `SUM(H1:H${row_num-1})`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];

        newRow = worksheet.addRow(ThirdLastRow);
        newRow.getCell(5).value = { formula: `SUM(E1:E${row_num-1})` }
        newRow.getCell(7).value ={ formula: `SUM(G1:G${row_num-1})` }
        newRow.getCell(8).value ={ formula: `SUM(H1:H${row_num-1})` }

        row_num =row_num +1

        SecondLastRow=[
                    "",
                    "",
                    "", 
                    "FL Sales Tax:", 
                    `E${row_num-1}*0.07`, 
                    "", 
                    `G${row_num-1}*0.07`, 
                    `H${row_num-1}*0.07`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];

        newRow = worksheet.addRow(SecondLastRow);
        newRow.getCell(5).value = { formula: `E${row_num-1}*0.07` }
        newRow.getCell(7).value ={ formula: `G${row_num-1}*0.07` }
        newRow.getCell(8).value ={ formula: `H${row_num-1}*0.07` }
        row_num =row_num +1

        LastRow=[
                    "",
                    "",
                    "",
                    "Total:", 
                    `E${row_num-2}+E${row_num-1}`, 
                    "",
                    `G${row_num-2}+G${row_num-1}`, 
                    `H${row_num-2}+H${row_num-1}`, 
                    "", 
                    "", 
                    "", 
                    ""
        ];
        newRow = worksheet.addRow(LastRow);
        newRow.getCell(5).value = { formula: `E${row_num-2}+E${row_num-1}` }
        newRow.getCell(7).value ={ formula: `G${row_num-2}+G${row_num-1}` }
        newRow.getCell(8).value ={ formula: `H${row_num-2}+H${row_num-1}` }





    }
    


    try {
        // Write the workbook to a buffer (Async operation)
        let buffer = await workbook.xlsx.writeBuffer(); 
        
        // Convert the buffer to a Blob object
        let blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        // Use FileSaver.js (saveAs) to trigger the download
        // Ensure FileSaver.js is loaded in your HTML
        saveAs(blob, filename);

        console.log("Excel file generated and downloaded successfully.");

    } catch (error) {
        console.error("Error generating or saving Excel file:", error);
        alert("Failed to generate Excel report.");
    }
});


$("input[name='choice']").on("click",function(){
    locationChoice=this.value
    console.log(locationChoice)
});


});

