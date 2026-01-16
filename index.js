let outstandingData=new Map();
let allData=[];
let locationChoice="NYC";
let locationTax=0.08875
let locationTaxString="NYC Sales Tax:"


async function generateReport(){
    icodesRan=[]
    if (allData.length == 0) {
        return alert("Please Upload All Items Excel Sheet");
    }

    if (outstandingData.size == 0) {
        return alert("Please Upload Outstanding Excel");
    }

    const filename = "onsite.xlsx";

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
    
        formatValueString=`FORMAT: -TRIPLE CHECK THAT ALL ITEMS ARE AS LISTED
    -Do a quick check—DRAPERIES/SHEERS, AREA RUG & light fixtures for all rooms
    —table lamps, nightstands & pillows/throws for bedrooms
    —bathrooms
    -Correct # counts
    -Correct vendor
    -FINAL changes/refreshes accounted for
    -Always login to vendors for pricing
    -Retail costs MUST include discounts
    Organic Modernism, WE, PB, C&B, CB2, William Sonoma, 
    -DO NOT guess name/pricing (UNSURE OF SOMETHING PLEASE HIGHLIGHT AND NOTE)
    (incorrect vendors create further issues for uplifts)
    -CORRECT FORMATTING IN Vendor + Item Name for ARTWORK
    (ARTWORK: “SIZE” TYPE) 
    -All sizing must be listed in Item Description column (rugs & art) and always FIRST 
    i.e. 36” round gold mirror
    -art format—IMG ART LOFT: size & type only
    i.e. 60” x 72” canvas
    `
    formatRow=[
        formatValueString,
        "",
        "",
        "",
        "", 
        ``, 
        "",
        ``, 
        ``, 
        "", 
        "", 
        "",
    ]



    let formatRowExcel=worksheet.insertRow(1,formatRow)
    formatRowExcel.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF999999' }
    }
    

    essentialString=
    `ESSENTIAL BEDDING INCLUDES: INSERT PILLOWS, INSERT DUVET, BASIC BEDDING SET 4 SHAMS, 1 FLAT SHEET, 1 FITTED SHEET
    TWIN- $600
    FULL- $800
    QUEEN- $850
    KING- $1,000																																
    `
    let essentialRow=[
        essentialString,
        "",
        "",
        "",
        "", 
        ``, 
        "",
        ``, 
        ``, 
        "", 
        "", 
        "",
    ]
    let essentialRowExcel=worksheet.insertRow(1,essentialRow)
    essentialRowExcel.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'b4cdcd' }
    }

    keyRowString=
    `KEY CHART: 
    COFFEE TABLE BOOKS- $65, COOKBOOKS- $45, KIDS BOOKS- $35, PROP BOOKS $26.99  - FLORALS: XL LARGE $24.99EA LARGE $18.99EA MED. $12.99EA SM $8.99EA CUT STEM $5.99EA BATHROOM: (1) BATH TOWEL $15EA (1) HAND TOWEL $10 WASHCLOTH $5EA 
    `
    keyRow=[
        keyRowString,
        "",
        "",
        "",
        "", 
        ``, 
        "",
        ``, 
        ``, 
        "", 
        "", 
        "",
    ]
    let keyRowExcel=worksheet.insertRow(1,keyRow)
    keyRowExcel.fill ={
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'b4cdcd' }
    }

        customRow=[
        "DO NOT TOUCH",
        "DO NOT TOUCH",
        "DESIGNER INPUT",
        "DESIGNER ADJUST ACCORDINGLY",
        "DESIGNER INPUT",
        "DO NOT TOUCH",
        "INVENTORY INPUT ONLY",
        "DO NOT TOUCH",
        "DO NOT TOUCH",
        "DESIGNER/INVENTORY INPUT",
        "DESIGNER/INVENTORY INPUT",
        "DESIGNER INPUT",
    ]
    let colorCodedRow=worksheet.insertRow(1,customRow)
    colorCodedRow.eachCell({ includeEmpty: true }, (cell) => {

        if (cell.value == "DO NOT TOUCH" || cell.value == "INVENTORY INPUT ONLY"){
            cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFA500'}
            };
        }
        else if (cell.value == "DESIGNER INPUT" || cell.value == "DESIGNER/INVENTORY INPUT"){
            cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00'}
            };
        }
        else if(cell.value == "DESIGNER ADJUST ACCORDINGLY"){
            cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFE0'}
            };
        }

    });

    blankspacefive=["","","","",""]

    darkgrey2="#999999"

    let template_row_five=worksheet.insertRow(1,blankspacefive)
    template_row_five.fill ={
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF999999' }
    }

    let template_row_four=worksheet.insertRow(1,["ADDRESS","","","",""])

    template_row_four.fill ={
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF999999' }
    }

    let template_row_three=worksheet.insertRow(1,["LUXURY FURNISHINGS INVENTORY","","","",""])
    template_row_three.fill ={
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF999999' }
    }

    let template_row_two=worksheet.insertRow(1,blankspacefive)
    template_row_two.fill ={
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF999999' }
    }

    let template_row_one=worksheet.insertRow(1,blankspacefive)
    template_row_one.fill ={
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF999999' }
    }


    worksheet.mergeCells('A1:L1')

    row2=worksheet.getRow(2)
    row2.alignment = { horizontal: 'center', vertical: 'middle'}
    row2.height=112.5

    worksheet.mergeCells('A2:L2')


    row3=worksheet.getRow(3)
    row3.alignment = { horizontal: 'center'}
    row3.alignment = { horizontal: 'center'}
    row3.font = {
    name: 'Montserrat',
    size: 18,
    color: { argb: 'FFFFFFFF' },
    };


    worksheet.mergeCells('A3:L3')

    row4=worksheet.getRow(4)
    row4.height=44.25
    row4.alignment = { horizontal: 'center'}
    row4.font = {
    name: 'Montserrat',
    size: 12,
    color: { argb: 'FFFFFFFF' },
    };
    worksheet.mergeCells('A4:L4')


    worksheet.mergeCells('A5:L5')

    // worksheet.mergeCells('A6:L6')

    row7=worksheet.getRow(7)
    row7.alignment = { horizontal: 'left', vertical: 'bottom',wrapText: true}
    row7.height=37.5;
    worksheet.mergeCells('A7:L7')

    row8=worksheet.getRow(8)
    row8.alignment = { horizontal: 'left', vertical: 'bottom',wrapText: true}
    row8.height=75;
    worksheet.mergeCells('A8:L8')

    row9=worksheet.getRow(9)
    row9.alignment = { horizontal: 'left', vertical: 'bottom',wrapText: true}
    row9.height=237;
    worksheet.mergeCells('A9:L9')


    worksheet.getRow(10).font = {
    name: 'Montserrat',
    size: 10,
    bold: true,
    };
    worksheet.getRow(10).height = 37.5;
    worksheet.getRow(10).alignment = { horizontal: 'center', vertical: 'middle'};





    let row_num =10; 
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
            totalFormula, unitCost, `${manufacturer}: ${description}`, "", 
            "" 
        ];


        if (icode == undefined) {
            let borderRow=worksheet.addRow(["", "", "", "", "", "", "", "", "", "", "", ""]); // Blank space 1
            borderRow.eachCell({ includeEmpty: true }, (cell) => {
                cell.border = {
                bottom: { style: 'thin' }
            };
            });
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
        else{
            newRow.font={
            size: 9,
            }
        }

        icodesRan.push(icode)
        row_num = row_num + 1

        // Set formulas in ExcelJS *after* adding the row, using the cell object
        if (icode != undefined) {

            estimateFormula = `C${newRow.number}*E${newRow.number}`;
            totalFormula = `C${newRow.number}*G${newRow.number}`;


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
        newRow.getCell(6).value = { formula: `SUM(F1:F${newRow.number-1})` }
        newRow.getCell(8).value ={ formula: `SUM(H1:H${newRow.number-1})` }
        newRow.getCell(9).value ={ formula: `SUM(I1:I${newRow.number-1})` }

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
        newRow.getCell(6).value = { formula: `F${newRow.number-1}*${locationTax}` }
        newRow.getCell(8).value ={ formula: `H${newRow.number-1}*${locationTax}` }
        newRow.getCell(9).value ={ formula: `I${newRow.number-1}*${locationTax}` }
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
        newRow.getCell(6).value = { formula: `F${newRow.number-2}+F${newRow.number-1}` }
        newRow.getCell(8).value ={ formula: `H${newRow.number-2}+H${newRow.number-1}` }
        newRow.getCell(9).value ={ formula: `I${newRow.number-2}+I${newRow.number-1}` }


    


    
    







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


};

document.addEventListener("DOMContentLoaded", function() {


$("label[for='switch_to_single']").on("click",function(){

    $(this).hide()
    $("label[for='switch_to_multi']").show()
    $("#singleFile").show()
    $("#multifile").hide()
});

$("label[for='switch_to_multi']").on("click",function(){

    $(this).hide()
    $("label[for='switch_to_single']").show()
    $("#singleFile").hide()
    $("#multifile").show()
});


 
const fileNameDisplay = document.getElementById('file-name-display');
$("#excel-file-input").on(("change"),function(){
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


const fileNameDisplayAsset=document.getElementById('file-name-display-asset');
$("#excel-file-input-outstanding").on(("change"),function(){

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

$("#process_files").on("click", async function() {
    generateReport()
});


$("input[name='choice']").on("click",function(){
    locationChoice=this.value
    console.log(locationChoice)

    if (locationChoice === "NYC"){
        locationChoice="NYC";
        locationTax=0.08875
        locationTaxString="NYC Sales Tax:"
        $("#option1-multi").prop("checked",true)
    }
    else if (locationChoice === "FL"){
        locationChoice="FL";
        locationTax=0.07
        locationTaxString="FL Sales Tax:"
        $("#option2-multi").prop("checked",true)
    }

});

$("input[name='choice-multi']").on("click",function(){
    locationChoice=this.value
    console.log(locationChoice)

    if (locationChoice === "NYC"){
        locationChoice="NYC";
        locationTax=0.08875
        locationTaxString="NYC Sales Tax:"
        $("#option1").prop("checked",true)
    }
    else if (locationChoice === "FL"){
        locationChoice="FL";
        locationTax=0.07
        locationTaxString="FL Sales Tax:"
        $("#option2").prop("checked",true)
    }

});


$("#excel-file-input-multi").on("change",function(){

    const files = event.target.files;

    if (files.length > 0) {
        allData=[]

        filenames=''
        
        for (let i = files.length-1; i>=0; i--) {
            currentFile=files[i]
            filenames=filenames + currentFile.name + " , "

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

            reader.readAsArrayBuffer(currentFile);
            

        }

        $("#file-name-display-all-multi").text(filenames)



    } else {

        $("#file-name-display-all-multi").text('No file selected');
    }



});


$("#excel-file-input-outstanding-multi").on("change",function(){

    const files = event.target.files;
    if (files.length > 0) {
        outstandingData=new Map()
        filenames=''
        for (let i = 0; i < files.length; i++) {
        const selectedFile = files[i];
        filenames=filenames + selectedFile.name +" , "

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
    }

    $("#file-name-display-asset-multi").text(filenames)

    } else {

        $("#file-name-display-asset-multi").text('No file selected')
    }
         


});


$("#process_files_multi").on("click",async function(){
    generateReport()
});



});

