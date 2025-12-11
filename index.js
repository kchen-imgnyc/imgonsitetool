let outstandingData=new Map()
let allData=[]

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

const generateReportButton =document.getElementById("process_files")

generateReportButton.addEventListener("click",function(){

    if (allData.length == 0){
        return alert("Please Upload All Items Excel Sheet")
    }

    if (outstandingData.size == 0){
        return alert("Please Upload Outstanding Excel")
    }

    filename="onsite.xlsx"

    header=["ICODE","QTY","ROOM + ITEM","ESTIMATE","ESTIMATE TOTAL","COST","TOTAL","PRICE","VENDOR + ITEM NAME","ITEM DESCRIPTION","IMAGE","REMOVAL NOTES"]
    blank_space=["","","","","","","","","","","",""]
    dataList=[header]

    for(let i=0;i<allData.length;i++){
        row_num = i + 2

        row=allData[i]
        icode=row[0]
        discription=row[1]
        unitCost=row[2]
        manufacturer=[3]

        extraData=outstandingData.get(icode)
        category=""
        imageID=""
        qty=0

        estimateFormula=`=B${row_num}*D${row_num}`
        totalFormula=`=B${row_num}*F${row_num}`
        imageLink="=IMAGE(\"https://imgnyc.rentalworks.cloud/api/v1/appimage/getimage?appimageid="+ String(imageID) +"&thumbnail=false\",4," +String(150)+","+ String(150)+")"


        if (extraData != undefined){
            if (manufacturer == "IMG"){
                manufacturer = "IMG ART LOFT:"
            }
            else if (manufacturer == "CUSTOM IMG"){
                manufacturer= "IMG CUSTOM"
            }
            else if ( 0 < manufacturer.length <=3){
                 manufacturer="IMG HOME EXCLUSIVE:"
            }
            category=extraData["Category"]
            imageID=extraData["ImageId"]
            qty=extraData["Quantity"]
        }
        newRow=[icode,qty,category,unitCost,estimateFormula,unitCost,totalFormula,unitCost,manufacturer + " - " + discription,imageLink,""]

        if (icode == undefined){
            dataList.push(blank_space)
            dataList.push(blank_space)
            newRow[1]=""
            newRow[2]=discription
            newRow[3]=""
            newRow[4]=""
            newRow[5]=""
            newRow[6]=""
            newRow[7] = ""
            newRow[8]=""
            newRow[10]=""
        }

        if (imageID == undefined || imageID == ""){
            newRow[10] = ""
        }

        dataList.push(newRow)

    }



    // Create a new workbook
    const wb = XLSX.utils.book_new();
    // Add a worksheet with the data
    const ws = XLSX.utils.aoa_to_sheet(dataList); // Use aoa_to_sheet for array of arrays
    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    // Write the file and trigger a client-side download
    XLSX.writeFile(wb, filename || 'export.xlsx');




});



});

