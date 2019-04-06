const xmlbuilder = require("xmlbuilder"),
    fs = require('fs'),
    path = require('path'),
    xml2js = require('xml2js'),
    Excel = require('exceljs')
var parseString = xml2js.parseString;
var listaProdNota = []


var workbook = new Excel.Workbook();
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Her';
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);
var worksheet = workbook.addWorksheet('PRODUTOS');

worksheet.columns = [
    { header: 'FORNEC', key: 'FORNEC', width: 10 },
    { header: 'RAZAO', key: 'RAZAO', width: 10 },
    { header: 'DESCRICAO', key: 'DESCRICAO', width: 10 },
    { header: 'CODIGO', key: 'CODIGO', width: 10 },
    { header: 'UNIDADE', key: 'UNIDADE', width: 10 },
    { header: 'NCM', key: 'NCM', width: 10 },
    { header: 'CEST', key: 'CEST', width: 10 },
    { header: 'QTD', key: 'QTD', width: 10 },
    { header: 'VUNIT', key: 'VUNIT', width: 10 },
    { header: 'ORIG', key: 'ORIG', width: 10 },
    { header: 'CFOP', key: 'CFOP', width: 10 }
];



fs.readdir('./xml', [{ withFileTypes: false }], function (err, files) {
    console.log(files)
    for (var file in files) {
        console.log(files[file])
        leituraarquivo(files[file])
    }
    workbook.xlsx.writeFile('EXCEL.xlsx')
        .then(function () {
            console.log(listaProdNota)

        });
});




function leituraarquivo(fileName) {
    fs.readFile(__dirname + '/xml/' + fileName, function (err, data) {
        // console.log(data)
        parseString(data, { explicitArray: false, ignoreAttrs: true }, function (err, result) {
            let prodNota = result.NFe.infNFe.det;
            if (!Array.isArray(prodNota)) {
                var matriz = [];
                matriz.push(prodNota);
                prodNota = matriz;
            }
            prodNota.forEach(function (item) {
                var TagICMS = Object.keys(item.imposto.ICMS)[0];
                var origem = 0;
                if (parseFloat(item.imposto.ICMS[TagICMS].orig) == 1) {
                    origem = 2
                };
                if (parseFloat(item.imposto.ICMS[TagICMS].orig) == 6) {
                    origem = 7
                };
                if (parseFloat(item.imposto.ICMS[TagICMS].orig) != 1 && parseFloat(item.imposto.ICMS[TagICMS].orig) != 6) {
                    origem = parseFloat(item.imposto.ICMS[TagICMS].orig)
                };
                var SITTRIB = '';
                SITTRIB += item.imposto.ICMS[TagICMS].CST | item.imposto.ICMS[TagICMS].CSOSN;
                listaProdNota.push({
                    'FORNEC': result.NFe.infNFe.emit.CNPJ,
                    'RAZAO': result.NFe.infNFe.emit.xNome,
                    'DESCRICAO': item.prod.xProd,
                    'CODIGO': item.prod.cProd,
                    'UNIDADE': item.prod.uCom,
                    'NCM': item.prod.NCM,
                    'CEST': item.prod.CEST,
                    'QTD': parseFloat(item.prod.qCom),
                    'VUNIT': parseFloat(item.prod.vUnCom),
                    'SITTRIB': SITTRIB,
                    'ORIG': origem,
                    'CFOP': item.prod.CFOP
                });
                worksheet.addRow({
                    'FORNEC': result.NFe.infNFe.emit.CNPJ,
                    'RAZAO': result.NFe.infNFe.emit.xNome,
                    'DESCRICAO': item.prod.xProd,
                    'CODIGO': item.prod.cProd,
                    'UNIDADE': item.prod.uCom,
                    'NCM': item.prod.NCM,
                    'CEST': item.prod.CEST,
                    'QTD': parseFloat(item.prod.qCom),
                    'VUNIT': parseFloat(item.prod.vUnCom),
                    'SITTRIB': SITTRIB,
                    'ORIG': origem,
                    'CFOP': item.prod.CFOP
                })
            });
        });
    });


}
