let outstandingData=new Map();
let allData=[];
let locationChoice="NYC";
let locationTax=0.08875
let locationTaxString="NYC Sales Tax:"

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

    const header = [
    "ICODE", "IMAGE", "QTY", "ROOM + ITEM", "ESTIMATE", "ESTIMATE TOTAL", 
    "COST", "TOTAL", "PRICE", "VENDOR + ITEM NAME", "ITEM DESCRIPTION", 
    "REMOVAL NOTES"
    ];

    worksheet.columns = [
        { header: header[0], key: 'icode', width: 12 },
        { header: header[1], key: 'image', width: 20 },
        { header: header[2], key: 'qty', width: 8 },
        { header: header[3], key: 'room_item', width: 30 },
        { 
            header: header[4], 
            key: 'estimate', 
            width: 15, 
            style: { numFmt: '$#,##0.00' } 
        },
        { 
            header: header[5], 
            key: 'estimate_total', 
            width: 18, 
            style: { numFmt: '$#,##0.00' } 
        },
        { 
            header: header[6], 
            key: 'cost', 
            width: 15, 
            style: { numFmt: '$#,##0.00' } 
        },
        { 
            header: header[7], 
            key: 'total', 
            width: 18, 
            style: { numFmt: '$#,##0.00' } 
        },
        { 
            header: header[8], 
            key: 'price', 
            width: 15, 
            style: { numFmt: '$#,##0.00' } 
        },
        { header: header[9], key: 'vendor_item', width: 35 },
        { header: header[10], key: 'description', width: 40 },
        { header: header[11], key: 'notes', width: 30 }
    ];

    // Apply universal styling (Font & Alignment) to all columns
    worksheet.columns.forEach(col => {
        col.style = col.style || {};
        col.style.font = { size: 10, name: 'Montserrat' };
        col.style.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    
    worksheet.getRow(1).font = {
    name: 'Montserrat',
    size: 10,
    bold: true,
    };
    
    worksheet.getRow(1).height = 37.5;
    worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle'};

    let row_num = 2; 
    for (let i = 0; i < allData.length; i++) {

        let row = allData[i];
        let icode = row[0];
        let description = row[1].toUpperCase();
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


        let estimateFormula = `C${row_num}*E${row_num}`;
        let totalFormula = `C${row_num}*G${row_num}`;

        let imageLink = '';

        if (manufacturer != undefined){
            manufacturer=manufacturer.toUpperCase()
            if (manufacturer == "IMG") {
                manufacturer = "IMG ART LOFT:";
            } else if (manufacturer == "CUSTOM IMG") {
                manufacturer = "IMG CUSTOM";
            } else if ( manufacturer.length > 0 && manufacturer.length <= 3) {
                manufacturer = "IMG HOME EXCLUSIVE:";
            }
        }
        else{
            manufacturer = "UNKNOWN"
        }

        if (extraData != undefined) {
            category = extraData["category"];
            imageID = extraData["imageID"];
            qty = extraData["qty"];
        }
        
        if (imageID != undefined) {
            imageLink = `IMAGE("https://imgnyc.rentalworks.cloud/api/v1/appimage/getimage?appimageid=${String(imageID)}&thumbnail=false",4,50,50)`;
        }


        let newRowValues = [
            icode,imageLink, qty, category, unitCost, estimateFormula, unitCost, 
            totalFormula, unitCost, `${manufacturer} - ${description}`, "", 
            "" 
        ];


        if (icode == undefined) {
            worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 1
            worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 2
            newRowValues = [
                "","", "", description, "", "", "", 
                "", "", "", "", 
                ""
            ];
            worksheet.getRow(row_num+1).height=10
            worksheet.getRow(row_num+2).height=10

            row_num = row_num + 2
        }


        // Handle missing ImageID
        if (imageID == undefined || imageID == "") {
            newRowValues[1] = "";
        }
        
        // Add the row
        let newRow = worksheet.addRow(newRowValues);

        if (icode == undefined){
            newRow.font={
            bold: true,
            }
        }

        icodesRan.push(icode)
        row_num = row_num + 1

        // Set formulas in ExcelJS *after* adding the row, using the cell object
        if (icode != undefined) {
            // Apply formula to ESTIMATE TOTAL (Column E, index 4)
            newRow.getCell(6).value = { formula: estimateFormula }; 
            
            // Apply formula to TOTAL (Column G, index 6)
            newRow.getCell(8).value = { formula: totalFormula }; 
            
            // Apply image formula to IMAGE (Column K, index 10)
            if (imageID != undefined) {
               newRow.getCell(2).value = { formula: imageLink };
            //    newRow.height = 112.5; // 150px is approx 112.5 points in Excel
            // newRow.height=75 //100px
                //newRow.height=56.25 //75px
                newRow.height=37.5 //50px

            }
        }
    }

    worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 1
    worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 2
    row_num = row_num + 2    

    //`B${row_num}*D${row_num}`

        ThirdLastRow=[
                    "",
                    "",
                    "",
                    "", 
                    "SubTotal:", 
                    ``, 
                    "", 
                    ``, 
                    ``, 
                    "", 
                    "", 
                    "", 
                    
        ];

        newRow = worksheet.addRow(ThirdLastRow);
        newRow.getCell(6).value = { formula: `SUM(F1:F${row_num-1})` }
        newRow.getCell(8).value ={ formula: `SUM(H1:H${row_num-1})` }
        newRow.getCell(9).value ={ formula: `SUM(I1:I${row_num-1})` }

        row_num =row_num +1

        SecondLastRow=[
                    "",
                    "",
                    "",
                    "", 
                    locationTaxString, 
                    ``, 
                    "", 
                    ``, 
                    ``, 
                    "", 
                    "", 
                    "", 
                
        ];

        newRow = worksheet.addRow(SecondLastRow);
        newRow.getCell(6).value = { formula: `F${row_num}*${locationTax}` }
        newRow.getCell(8).value ={ formula: `H${row_num}*${locationTax}` }
        newRow.getCell(9).value ={ formula: `I${row_num}*${locationTax}` }
        row_num =row_num +1

        LastRow=[
                    "",
                    "",
                    "",
                    "",
                    "Total:", 
                    ``, 
                    "",
                    ``, 
                    ``, 
                    "", 
                    "", 
                    "", 
        
        ];
        newRow = worksheet.addRow(LastRow);
        newRow.getCell(6).value = { formula: `F${row_num-1}+F${row_num}` }
        newRow.getCell(8).value ={ formula: `H${row_num-1}+H${row_num}` }
        newRow.getCell(9).value ={ formula: `I${row_num-1}+I${row_num}` }

        const sheet2 = workbook.addWorksheet('Client')

        blackspacefive=["","","","",""]

        darkgrey2="#999999"

        let template_row_one=sheet2.addRow(blackspacefive)
        template_row_one.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
        }
        sheet2.mergeCells('A1:Z1')

        let template_row_two=sheet2.addRow(blackspacefive)
        template_row_two.alignment = { horizontal: 'center', vertical: 'middle'}
        template_row_two.height=112.5
        template_row_two.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
        }
        sheet2.mergeCells('A2:Z2')
        
        let template_row_three=sheet2.addRow(["LUXURY FURNISHINGS INVENTORY","","","",""])
        template_row_three.alignment = { horizontal: 'center'}
        template_row_three.font = {
        name: 'Montserrat',
        size: 18,
        color: { argb: 'FFFFFFFF' },
        };
        template_row_three.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
        }
        sheet2.mergeCells("A3:Z3")

        let template_row_four=sheet2.addRow(["ADDRESS","","","",""])
        template_row_four.height=44.25
        template_row_four.alignment = { horizontal: 'center'}
        template_row_four.font = {
        name: 'Montserrat',
        size: 12,
        color: { argb: 'FFFFFFFF' },
        };
        template_row_four.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
        }
        sheet2.mergeCells("A4:Z4")

        let template_row_five=sheet2.addRow(blackspacefive)
        template_row_five.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
        }
        sheet2.mergeCells('A5:Z5')

        // let colA=sheet2.getColumn(1)
        // colA.width=62.25

    // const img = document.getElementById('logo');
    // const canvas = document.createElement('canvas');
    
    // canvas.width = img.naturalWidth;
    // canvas.height = img.naturalHeight;
    
    // // Draw the image onto the canvas
    // const ctx = canvas.getContext('2d');
    // ctx.drawImage(img, 0, 0);
    
    // const dataURL = canvas.toDataURL('image/jpeg'); // Specify the MIME type
    
    // const imageId = workbook.addImage({
    // base64: dataURL,
    // extension: 'png',
    // });

    // sheet2.addImage(imageId, {
    //     tl: { col: 0, row: 1 },
    //     ext: { width: img.naturalWidth, height: img.naturalHeight }
    // });
        



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

    if (locationChoice === "NYC"){
        locationChoice="NYC";
        locationTax=0.08875
        locationTaxString="NYC Sales Tax:"
    }
    else if (locationChoice === "FL"){
        locationChoice="FL";
        locationTax=0.07
        locationTaxString="FL Sales Tax:"
    }

});


});

