var XlsxTemplate = require('xlsx-template')
var fs = require('fs');
var path = require('path');

fs.readFile(path.join(__dirname, 'template', 'mau.xlsx'), function(err, data){
    var template = new XlsxTemplate(data)

    var sheetNumber = 1;

    var values = {
        ngay: '04',
        thang: '04',
        nam: '2020',
        people: [
            {stt:1, name:'Nguoi thu 1'},
            {stt:2, name:'Nguoi thu 2'},
            {stt:3, name:'Nguoi thu 3'},
            {stt:4, name:'Nguoi thu 4'},
        ]
    }

    template.substitute(sheetNumber, values);

    var output = template.generate();

    // save file
    fs.writeFileSync(path.join(__dirname, 'output', 'out.xlsx'), output, 'binary');
})