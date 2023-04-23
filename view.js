let $ = require('jquery') // Module jquery to select
let fs = require('fs') // Module fs to rw file

const reader =  require('xlsx') // Module xlsx
// For callback parameter 
const file = reader.readFile('./test.xlsx')

let worksheets = {} 



$('#submit').on('click', () => {
        
        let res = confirm("�A�T�w�n�e�X�H�U����ƶ�?\n" + 
                            document.getElementById("name").value + '\n' +
                            document.getElementById("id").value + '\n' + 
                            document.getElementById("dep").value + '\n' + 
                            document.getElementById("phone").value + '\n' + 
                            document.getElementById("date").value + '\n' )

        let total = parseInt(document.getElementById("o-19w").value) + 
                    parseInt(document.getElementById("o-16w").value) +
                    parseInt(document.getElementById("o-19m").value) + 
                    parseInt(document.getElementById("o-16m").value) + 
                    parseInt(document.getElementById("w10-32").value) +  
                    parseInt(document.getElementById("w10-64").value) +  
                    parseInt(document.getElementById("sas93").value) + 
                    parseInt(document.getElementById("sas-94").value) +  
                    parseInt(document.getElementById("vs-15").value) +
                    parseInt(document.getElementById("vs-13").value) +
                    parseInt(document.getElementById("vs-12").value) +
                    parseInt(document.getElementById("ev").value) + 
                    parseInt(document.getElementById("nat").value) + 
                    parseInt(document.getElementById("wu").value) +
                    parseInt(document.getElementById("usb").value) + 
                    parseInt(document.getElementById("pho").value) + 
                    parseInt(document.getElementById("pro").value) 
        
        if ( total > 0 && valid() ) {
            worksheets = {} 
            for (const sheetName of file.SheetNames) {
                worksheets[sheetName] = reader.utils.sheet_to_json(file.Sheets[sheetName])
            }
            let temp = worksheets["Sheet1"]
            let sno = 0
            if(temp.length == 0 ) sno = 0 
            else {
                sno = temp[temp.length - 1]["�y��"] + 1
            }
            worksheets.Sheet1.push({
                "�y��" : sno,
                "�m�W" : document.getElementById("name").value, 
                "���" : document.getElementById("date").value, 
                "�Ǹ�" : document.getElementById("id").value, 
                "���" : document.getElementById("dep").value,
                "�q��" : document.getElementById("phone").value, 
                "�`��" : total,
                "(32)Windows10" : document.getElementById("w10-32").value,
                "(64)Windows10" : document.getElementById("w10-64").value,
                "Office2016" : document.getElementById("o-16w").value,
                "Office2019" : document.getElementById("o-19w").value,
                "Office2016Mac" : document.getElementById("o-16m").value, 
                "Office2019Mac" : document.getElementById("o-19m").value,
                "SAS9.3" : document.getElementById("sas93").value,
                "SAS9.4" : document.getElementById("sas-94").value,
                "VisualStudio2012" : document.getElementById("vs-12").value, 
                "VisualStudio2013" : document.getElementById("vs-13").value, 
                "VisualStudio2015" : document.getElementById("vs-15").value, 
                "EVIEWS" : document.getElementById("ev").value, 
                "�۵M��J�k" : document.getElementById("nat").value, 
                "�L����" : document.getElementById("wu").value, 
                "����333" : document.getElementById("usb").value,
                "PhotoExplorer" : document.getElementById("pho").value, 
                "ProtelDxp" : document.getElementById("pro").value
            })

            reader.utils.sheet_add_json(file.Sheets["Sheet1"], worksheets.Sheet1)
            reader.writeFile(file,'./test.xlsx')
        }

        else {
            var form = document.getElementById("submit")
            form.addEventListener('submit', stopSubmit)
        }
})

function valid() {
    if (document.getElementById("name").val == "" || document.getElementById("date").val == "" || 
        document.getElementById("id").val == "" || document.getElementById("dep").val == "" || 
        document.getElementById("phone").val == "")
       return false
    return true 
}

function stopSubmit(event) {
    event.preventDefault(); 
}

function process() {
    
}