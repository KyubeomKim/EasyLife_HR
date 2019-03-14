var express = require('express');
var router = express.Router();

const fs = require('fs');
const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
var filename = "Data Analytics,Engineering.xlsx";
var personIndex = 2;

/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('index', { title: 'Express' });
});

router.get('/scoring', function(req, res, next) {
    let workbook = XLSX.readFile("./data/" + filename)
    let worksheetRaw = workbook.Sheets["Raw"];
    // let worksheetTotal = workbook.Sheets["total"];

    // console.log(worksheetRaw);
    var questionLength = worksheetRaw["!ref"].split(":")[1][0].charCodeAt(0) - 68;
    var personLength = parseInt(worksheetRaw["!ref"].split(":")[1][1]);

    for (var j = 0; j < questionLength; j++) {
        console.log(String.fromCharCode(69 + j) + personIndex)
        console.log(worksheetRaw[String.fromCharCode(69 + j) + personIndex].v)
    }
    personIndex += 1;
    res.render('index', { title: 'next' });


    // for (var i = 1; i <= personLength; i++) {
    //     for (var j = 0; j < questionLength; j++) {
    //         console.log(worksheetRaw[String.fromCharCode(69 + j) + i].v)
    //     }
    // }

    // console.log(worksheetRaw["D1"]);

    // var totalProfitList = []
    // for (var i = 2; i < 5; i++) {
    //     totalProfitList.push(worksheetTotal["C" + i].v + worksheetDashboard["E" + i].v);
    // }
    // console.log(XLSX.utils.sheet_to_json(worksheetDashboard))

    // res.render('dashboard', { params: XLSX.utils.sheet_to_json(worksheetDashboard), totalProfitList: totalProfitList });
});

router.get('/addIndex', function(req, res, next) {
    personIndex += 1;
    res.render('index', { title: 'next' });
})

module.exports = router;