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

            dataObj={
                "discription":discription,
                "unitCost":unitCost,
                "manufacturer":manufacturer
            }
            
            
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






});



});

