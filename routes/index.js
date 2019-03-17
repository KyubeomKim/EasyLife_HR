var express = require('express');
var router = express.Router();

const fs = require('fs');
const XLSX = require("xlsx");
const XLSX_CALC = require('xlsx-calc');
var formulajs = require("formulajs");
var config = JSON.parse(fs.readFileSync("./config/config.json", "utf8"))
var interviewer;

/* GET home page. */
router.get('/login', function(req, res, next) {
    res.render('login', { title: config.title });
});

router.post('/login', function(req, res, next) {
    var currentInterviewer = config.interviewer.find(function(element) {
        return element.id == req.body.id && element.birthday == req.body.birthday;
    })
    if (currentInterviewer == undefined) {
        console.log("사용자가 없습니다")
        res.redirect('/login');
    } else {
        console.log(currentInterviewer)
        interviewer = currentInterviewer;
        res.redirect('/scoring');
    }
});

router.get('/scoring', function(req, res, next) {
    if(interviewer == undefined) {
        res.redirect('/login');
        // res.render('scoring');
    } else {
        var workbooks = [];
        interviewer.scoring_type.forEach(function(element) {
            var workbook = XLSX.readFile("./data/" + element + ".xlsx").Sheets["Raw"]
            workbooks.push(workbook)
        })

        var params = [];
        if(interviewer.scoring_item == "all") {
            config.scoring.forEach(function (scoringElement) {
                var object = {};
                var data = [];
                object.label = scoringElement.label;
                scoringElement.data_to_display.forEach(function(dataElement) {
                    var obj = {};
                    obj["question"] = workbooks[0][dataElement + "1"].v
                    obj["answer"] = workbooks[0][dataElement + "2"].v
                    console.log(obj)
                    data.push(obj)
                })
                object.data = data; 
                params.push(object);
            })
        }
        console.log(params)

        res.render('scoring', { params: params, interviewer: interviewer });
    }

    // // console.log(worksheetRaw);
    // var questionLength = worksheetRaw["!ref"].split(":")[1][0].charCodeAt(0) - 68;
    // var personLength = parseInt(worksheetRaw["!ref"].split(":")[1][1]);

    // for (var j = 0; j < questionLength; j++) {
    //     console.log(String.fromCharCode(69 + j) + personIndex)
    //     console.log(worksheetRaw[String.fromCharCode(69 + j) + personIndex].v)
    // }
    // personIndex += 1;
    


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