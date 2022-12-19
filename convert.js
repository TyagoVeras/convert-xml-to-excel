import fs from 'node:fs';
import excelJs from 'exceljs';
import xml2json from 'xml2json';

function createFileExcel(input, output) {
  const fileExcel = new excelJs.stream.xlsx.WorkbookWriter({
    stream: fs.createWriteStream(`${output}`)
  }); 

  const worksheetFile = fileExcel.addWorksheet('Products');
  worksheetFile.columns = [
      { header: 'CODIGO', key: 'code', width: 10 },
      { header: 'CODIGO DE BARRA', key: 'bar_code', width: 10 },
      { header: 'DESCRICAO', key: 'description', width: 30 },
      { header: 'QUANTIDADE', key: 'qntd', width: 30 },
      { header: 'PRECO UNITARIO', key: 'price_unit', width: 10 },
      { header: 'PRECO TOTAL', key: 'price_total', width: 10 },
    ];

    const makeProduct = (product) => {
      return {
        code: product.cProd,
        bar_code: product.cEAN,
        descricao: product.xProd,
        qntd: product.qCom,
        price_unit: product.vUnCom,
        price_total: product.vProd,
      }
    }
     fs.readFile(input, (err, data) => {
      const note = xml2json.toJson(data, {object: true});
      const products = note.nfeProc.NFe.infNFe.det
      products.forEach((product) => {
        worksheetFile.addRow(makeProduct(product.prod)).commit();
      });
      fileExcel.commit();
    });

  }


createFileExcel('./input.xml', './result.xlsx')