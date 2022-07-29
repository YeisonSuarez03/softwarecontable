import React, { useEffect, useState } from "react";
import ExcelFile from "react-export-excel/dist/ExcelPlugin/components/ExcelFile";
import ExcelColumn from "react-export-excel/dist/ExcelPlugin/elements/ExcelColumn";
import ExcelSheet from "react-export-excel/dist/ExcelPlugin/elements/ExcelSheet";
import { OutTable, ExcelRenderer } from "react-excel-renderer";
import * as XLSX from "xlsx";
import "./App.css";
import { isEmpty } from "xlsx-populate/lib/xmlq";

function App() {

  const [dataESF, setDataESF] = useState(null);
  const [dataERI, setDataERI] = useState(null);
  const [dataFormattedESF, setDataFormattedESF] = useState(null);
  const [dataFormattedERI, setDataFormattedERI] = useState(null);
  
  const [cols, setCols] = useState([]);
  const [rows, setRows] = useState([]);

  const [headerInfo, setHeaderInfo] = useState(null);
  const [footerInfo, setFooterInfo] = useState([
    ["ANDREA GOMEZ VARON" ,"MARIA FERNANDA GALVIS R.", "YALILA ROJAS"],
    ["Representante Legal", "Contador Publico T.P. 137429-T", "Revisor Fiscal T.P. 24545-T"]
  ]);

  const labelsByCuenta = {
    "4170": "Cuotas de Administración"
  }

  const getExcelDataAndFormat = (rows, headerDoc = true) => {
    //items con codigo de 1 digito seran las propiedades principales, sus secundarias seran
    //codigos con 2 digitos. Los valores seran items con codigo de 4 digitos
    let dataObject = {
      "esf": {},
      "eri": {}
    }
    let parentPropertyFirstLevel = "";
    let parentPropertySecondLevel = "";
    console.log(rows)
    //
    let rowsByFormat = docformat == "csv" ? rows.map(values => [...values[0].replace(/[\",]/g, "").split(/\u0000/)]) : rows
    console.log(rowsByFormat)
    console.log("docformat: ", docformat)
    headerDoc && setHeaderInfo(
      docformat == "csv" ? [
        rowsByFormat[0].filter(v => v !== '' && v !== '\"'), 
        rowsByFormat[1].filter(v => v !== '' && v !== '\"'), 
        rowsByFormat[2].filter(v => v !== '' && v !== '\"')
      ]
      : [
        rowsByFormat[0],
        rowsByFormat[1],
      rowsByFormat[2],
    ]
      )
    rowsByFormat.forEach(values => {
      values.forEach(value => {
        if (typeof value == "string" && value !== "") {
          let textSplit = value.split("      ")
          if (!isNaN(textSplit[0]) && textSplit[0] != "0.00" ) {
            let dataObjectSheet = textSplit[0].at(0) > 3 ? "eri" : "esf"
            console.log("textSplit[0]: ", textSplit[0], textSplit[0].at(0))
            console.log("dataObjectSheet: ", dataObjectSheet, dataObject[dataObjectSheet])
            console.log(value)
            let saldoInicialPositionByIndex = docformat === "csv" 
            ?  parseInt(values.length.toString().at(-1))+1
            : 6
            switch (textSplit[0].length) {
              case 1:
                parentPropertyFirstLevel = textSplit[1]
                Object.defineProperty(
                  dataObject[dataObjectSheet], textSplit[1], 
                  {value: {}, writable: true, enumerable: true}) 
                
                break;
              
              case 2:
                parentPropertySecondLevel = textSplit[1]
                Object.defineProperty(
                  dataObject[dataObjectSheet][parentPropertyFirstLevel], textSplit[1], 
                  {value: [], writable: true, enumerable: true}) 

                break;
              
              case 4:
                let arrayTextSplit = [textSplit[0], labelsByCuenta[textSplit[0]] || textSplit[1]].join("      ");
                console.log("saldoInicialPositionByIndex0: ", textSplit[0] === "1345" ? saldoInicialPositionByIndex :"")
                console.log("saldoInicialPositionByIndex1: ", textSplit[0] === "1345" ? Math.abs(parseFloat(rowsByFormat.find(v => !isEmpty({children: v[0]}) && v[0].split("      ")[0] === "13459590").at(saldoInicialPositionByIndex))) :"")
                console.log("saldoInicialPositionByIndex2: ", textSplit[0] === "1345" ? Math.abs(parseFloat(values.at(saldoInicialPositionByIndex))) :"")
                let arrayValues = textSplit[0].at(0) > 3 

                //values para ERI
                ? [
                  arrayTextSplit, 
                  (+Math.abs(parseFloat(values.at(-1))) - Math.abs(parseFloat(values[saldoInicialPositionByIndex]))).toString(),
                  "0",
                  values.at(-1).toString()
                ]
                //values para ESF
                : [
                  arrayTextSplit, 
                  textSplit[0] === "1345" 
                    ? (Math.abs(parseFloat(rowsByFormat.find(v => !isEmpty({children: v[0]}) && v[0].split("      ")[0] === "13459590").at(-1))) - Math.abs(parseFloat(values.at(-1)))).toString()
                    : values.at(-1).toString(), 
                  textSplit[0] === "1345" 
                    ?  (Math.abs(parseFloat(rowsByFormat.find(v => !isEmpty({children: v[0]}) && v[0].split("      ")[0] === "13459590").at(saldoInicialPositionByIndex))) - Math.abs(parseFloat(values.at(saldoInicialPositionByIndex)))).toString()
                    : values[saldoInicialPositionByIndex].toString(), 
                  
                ]
                dataObject[dataObjectSheet][parentPropertyFirstLevel][parentPropertySecondLevel].push(arrayValues)

                break;
            
              default:
                if (
                  textSplit[0].slice(0,4) === "1110" ||
                  textSplit[0].slice(0,4) === "1120"
                  ) {
                  let arrayValues = [textSplit.join("      "), values[saldoInicialPositionByIndex].toString(), values[values.length-1].toString()]
                  dataObject[dataObjectSheet][parentPropertyFirstLevel][parentPropertySecondLevel].push(arrayValues)
                }
                break;
            }
          }
          console.log(value.split("      ")[0])
        }
      })
    })
    console.log("dataObject: ", dataObject)
    console.log("dataESF: ", dataObject.esf)
    console.log("dataERI: ", dataObject.eri)
    setDataESF(dataObject.esf);
    setDataERI(dataObject.eri);
  }

  const getVariationsAndTotalsESF = (data, isEsf = true) => {
    let newData = data;
    let keys = Object.keys(newData);
    Object.values(newData).forEach((v, i) => {
      let values = Object.values(v);
      values.forEach((value) => {
        let total1 = 0;
        let total2 = 0;
        let total4 = 0;
        let total3Values = [];
        for (let i = 0; i < value.length; i++) {
          let variationValue = Math.abs(
            +Math.abs(parseFloat(value[i][1])) -
              Math.abs(parseFloat(value[i][2]))
          ).toString();
          console.log(parseFloat(value[i][1]));
          console.log(parseFloat(value[i][2]));
          console.log("value: ",value[i][0])
          let splitCode = value[i][0].split("      ")[0]
          total1 += splitCode.length === 4 ? parseFloat(value[i][1]) : 0;
          total2 += splitCode.length === 4 ? parseFloat(value[i][2]) : 0;
          total4 += !isEsf && splitCode.length === 4 ? parseFloat(value[i][3]) : 0;
          isEsf ? value[i].push(variationValue) : value[i].splice(3,0,variationValue); 
          splitCode.length === 4 && total3Values.push(parseFloat(variationValue));
        }
        let allTotalValues = [
          " ",
          total1.toString(),
          total2.toString(),
          total3Values.reduce((prev, v) => prev + v).toString(),
          !isEsf ? total4.toString() : " "
        ];
        value.push(allTotalValues);
        console.log(allTotalValues);
      });
      let allTotalValuesFiltered = values
        .reduce((prev, v) => prev.concat(v))
        .filter((v) => v[0] === " ")
        .reduce((prev, v) => [
          `TOTAL ${keys[i].toLocaleUpperCase()}`,
          Math.abs(parseFloat(prev[1]) + parseFloat(v[1])).toString(),
          Math.abs(parseFloat(prev[2]) + parseFloat(v[2])).toString(),
          Math.abs(parseFloat(prev[3]) + parseFloat(v[3])).toString(),
          !isEsf ? Math.abs(parseFloat(prev[4]) + parseFloat(v[4])).toString() : " ",
        ]);
      newData[keys[i]].total = [allTotalValuesFiltered];
    });
    //generamos total de PASIVO + PATRIMONIO
    if (Object.keys(newData).includes("PASIVO") && Object.keys(newData).includes("PATRIMONIO")) {
      newData["PATRIMONIO"].total.push([
        "TOTAL PASIVO MAS PATRIMONIO",
        (Math.abs(parseFloat(newData["PASIVO"].total[0][1])) + Math.abs(parseFloat(newData["PATRIMONIO"].total[0][1]))).toString(),
        (Math.abs(parseFloat(newData["PASIVO"].total[0][1])) + Math.abs(parseFloat(newData["PATRIMONIO"].total[0][1]))).toString(),
        (Math.abs(parseFloat(newData["PASIVO"].total[0][1])) + Math.abs(parseFloat(newData["PATRIMONIO"].total[0][1]))).toString(),
      ])
    }
    console.log("newData: ", newData);

    return newData
  };

  const formatDataToCell = (v) => {
    console.log(v)
    console.log(typeof v[0])
    return v.map((data, i) => ({
      value: data != "0" 
      ? !isNaN(data.split("      ").at(-1)) && Math.sign(parseFloat(data.split("      ").at(-1))) == -1
        ?  `(${Math.abs(data.split("      ").at(-1))})` 
        :  data.split("      ").at(-1)
      : "-",
      style:
        typeof v[0] !== 'undefined' && v[0].includes("TOTAL") && isNaN(data)
          ? { font: titleFontStyles }
          : v[0].includes("TOTAL")
          ? { ...totalNumberStyle }
          : v[0] === " "
          ? { ...titleNumberStyle }
          : isNaN(data)
          ? v[i + 1] === " "
            ? { font: { ...basicFontStyles, bold: true } }
            : { font: basicFontStyles }
          : { ...basicNumberStyle },
    }));
  }

  const mapDataToExcelFormat = (data) => {
    let newData = data;
    let newDataFormatted = [];
    console.log(Object.keys(newData));
    Object.values(newData).forEach((v, i) => {
      let values = Object.values(v);
      console.log("title: ", Object.keys(newData)[i]);
      let titleSection = [
        {
          value: Object.keys(newData)[i] ,
          style: { font: titleFontStyles },
        },
        " ",
        " ",
        " ",
      ];
      let titles = Object.keys(v).map((data) => {
        console.log("data: ", data);
        return [
          {
            value: data !== "total" ? data : " ",
            style: isNaN(data)
              ? { font: { ...basicFontStyles, bold: true } }
              : { ...basicNumberStyle },
          },
          " ",
          " ",
          " ",
        ];
      });
      console.log(titles);
      let mapDataFormatted = values.map((value) => {
        return value.map((v) => {
          console.log("pop V:",v)
          v.at(-1) === " " && v.pop();
          return formatDataToCell(v)
        });
      });
      let titleAndDataFormatted = mapDataFormatted.map((v, i) => {
        return [titles[i], ...v];
      });
      let finalDataFormatted = [
        titleSection,
        ...titleAndDataFormatted.reduce((prev, v) => prev.concat(v)),
      ];
      console.log(mapDataFormatted);
      console.log(titleAndDataFormatted);
      console.log(finalDataFormatted);
      newDataFormatted.push(...finalDataFormatted);
    });
    console.log(newDataFormatted);
    // let headerInfoFormatted = headerInfo.map(v => formatDataToCell(v))
    // console.log("headerInfoFormatted: ", headerInfoFormatted)
    return newDataFormatted
};

  const formatHeaderInfo = (headerInfo, title) => {
    return headerInfo.map((v, i) => ([{value: v.at(-1).replace("Balance de Prueba", title, "gi"), style: {font: titleFontStyles}}]))
  }

  const titleFontStyles = {
    name: "Century Gothic",
    sz: "10",
    bold: true,
  };

  const basicFontStyles = {
    name: "Century Gothic",
    sz: "10",
  };
  const centerAlignment = {
    vertical: "center",
    horizontal: "center",
    wrapText: false,
  };

  const rightAlignment = {
    vertical: "right",
    horizontal: "right",
    wrapText: false,
  };

  const basicNumberStyle = {
    font: { ...basicFontStyles },
    numFmt: "$#,###.00",
    alignment: rightAlignment,
  };

  const titleNumberStyle = {
    font: {
      ...titleFontStyles,
    },
    numFmt: "$#,###.00",
    alignment: rightAlignment,
    border: { top: { style: "thin", color: { rgb: "000000" } } },
  };

  const totalNumberStyle = {
    font: {
      ...titleFontStyles,
    },
    numFmt: "$#,###.00",
    alignment: rightAlignment,
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "medium", color: { rgb: "000000" } },
    },
  };

  const [docformat, setDocformat] = useState(null)

  const fileHandler = (event) => {
    let fileObj = event.target.files[0];
    console.log(event.target.files[0]);
    setDocformat(event.target.files[0].name.split(".").at(-1))
    //just pass the fileObj as parameter
    ExcelRenderer(fileObj, (err, resp) => {
      if (err) {
        console.log(err);
      } else {
        setCols(resp.cols);
        setRows(resp.rows);
      }
    });
  };

  useEffect(() => {
    if (rows !== null && rows.length > 0) {
      getExcelDataAndFormat(rows);
    }
  }, [rows])

  useEffect(() => {
    if (dataESF !== null && 
      dataERI !== null) {
      let newDataESF = getVariationsAndTotalsESF(dataESF);
      let newDataERI = getVariationsAndTotalsESF(dataERI, false);
      setDataESF(newDataESF)
      setDataERI(newDataERI)
    }
  }, [dataESF, dataERI]);

  useEffect(() => {
    if (
      (dataESF !== "" && dataESF !== null) &&
      (dataERI !== "" && dataERI !== null)
    ) {
      let newDataESFFormatted = mapDataToExcelFormat(dataESF);
      let newDataERIFormatted = mapDataToExcelFormat(dataERI);
      console.log("headerInfo: ", headerInfo)
      setDataFormattedESF([
        {columns: [" "], data: [...formatHeaderInfo(headerInfo, "Estado de Situación Financiera"), [" "]]}, 
        {columns: ["Items", "Mayo 2022", "Abril 2022", "Variación"], data: newDataESFFormatted},
        {columns: [" ", " ", " ", " "], data: [[" "], ...footerInfo.map(values=> {return values.map(v => ({value: v, style: {font: titleFontStyles}}))})]}
      ]);
      setDataFormattedERI([
        {columns: [" "], data: [...formatHeaderInfo(headerInfo, "Estado de Resultado Integral"), [" "]]},
        {columns: ["Items", "Mayo 2022", "Abril 2022", "Variación", "Acumulado Mayo 2022"], data: newDataERIFormatted},
        {columns: [" ", " ", " ", " "], data: [[" "], ...footerInfo.map(values=> {return values.map(v => ({value: v, style: {font: titleFontStyles}}))})]}
      ])
    }
  }, [dataESF, dataERI]);
  

  return (
    <div className="App">
      <div>
        <input type="file" onChange={fileHandler} style={{ padding: "10px" }} />
        {/* <div>
          <OutTable
            data={rows}
            columns={cols}
            tableClassName="ExcelTable2007"
            tableHeaderRowClass="heading"
          />
        </div> */}
      </div>

      {console.log("dataFormattedESF: ", dataFormattedESF)}
      {dataFormattedESF !== null && (
        <ExcelFile filename="Estados Financieros">
          <ExcelSheet
            dataSet={dataFormattedESF}
            name="ESTADO DE SITUACION FINANCIERA"
          >
          </ExcelSheet>
          <ExcelSheet
            dataSet={dataFormattedERI}
            name="ESTADO DE RESULTADO INTEGRAL"
          >
          </ExcelSheet>
        </ExcelFile>
      )}
    </div>
  );
}

export default App;
