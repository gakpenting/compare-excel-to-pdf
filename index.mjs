import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import pdf from './pdf-parse/index.js';

const workbook = XLSX.readFile('./example_excel.xlsx');
const __dirname = path.resolve();
let to_json =async function to_json(workbook) {
    var result = {};
    const sementara=[];

    // console.log(workbook.SheetNames)
    // change sheet to json
    let jsonParse = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"], {header:1});
    // change remove header
    let removeHeader=jsonParse.splice(1,jsonParse.length)
        for(let b of removeHeader){
sementara.push({name:b[0],pdf_file:null})
    }
    // read example_pdf directory
                 const files=fs.readdirSync('./example_pdf')
             // read pdf file one by one and push it to semuaFile array
             let semuaFile=[];
                for await(let c of files){
                    var absolute_path_to_pdf = path.join(__dirname, 'example_pdf',c)
                        let dataBuffer = fs.readFileSync(absolute_path_to_pdf);
                        const z=await pdf(dataBuffer)
                        const j=z.text.trim()
                        semuaFile.push({name:c.toLowerCase(),text:j.split('\n').join(' ').toLowerCase()})
                }
                //compare excel data to pdf file contents and name
                for(let c of sementara){
                    for(let d of semuaFile){
                        if(d.name.toLowerCase().indexOf(c.name.toLowerCase())!==-1||d.text.toLowerCase().indexOf(c.name.toLowerCase())!==-1){
if(c.pdf_file){
    c.pdf_file.push(d.name)
}else{
    c.pdf_file=[]
    c.pdf_file.push(d.name)
}
                        }
                    }
                }
                let stringifySementara=sementara.map(a=>{
                return  {name:a.name,pdf_file:JSON.stringify(a.pdf_file)}
                })
    //create a new workbook
    let wb = XLSX.utils.book_new();
    //change json to sheet
    let ws=XLSX.utils.json_to_sheet(stringifySementara);
    //create a new workbook named compared_pdf
    XLSX.utils.book_append_sheet(wb, ws, "compared_pdf");
    /* generate an XLSX file */
    XLSX.writeFile(wb, "compared_pdf.xlsx");

};
to_json(workbook)