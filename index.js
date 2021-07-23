const firebaseHelper = require('./helpers/FirebaseHelper')
const fs = require('fs');
const request = require('request');
const cmd = require('node-cmd');
const express = require("express");
const app = express();
const PORT = process.env.PORT || 8080;
const http = require('https');
const robot = require("robotjs");

const EXCEL_EXE = " \"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE\""

require('core-js/modules/es.promise');
require('core-js/modules/es.string.includes');
require('core-js/modules/es.object.assign');
require('core-js/modules/es.object.keys');
require('core-js/modules/es.symbol');
require('core-js/modules/es.symbol.async-iterator');
require('regenerator-runtime/runtime');

const Excel = require('exceljs/dist/es5');
const uuid = require("uuid-v4");


app.listen(PORT, () => console.log(`Listing in port ${PORT}`));

firebaseHelper.admin.database().ref("listener_documents").child("document").on('value', (snapshot) => {
    if (snapshot.exists()) {

        let data = snapshot.val()

        let path = `${data.id}.xlsx`
        let command = `${EXCEL_EXE} "${__dirname}/${path}"`

        console.log(path)
        console.log(command)
        console.log(__dirname)

        download(data.url, path, () => {
            cmd.run(command, (err, data, stderr) => {
                }
            );

            setTimeout(() => {
                robot.mouseClick();
                robot.mouseClick();
                robot.keyTap("left");
                robot.keyTap("enter");
                setTimeout(() => {
                    robot.keyTap("enter");

                    setTimeout(() => {
                        robot.keyTap("enter");

                        setTimeout(() => {


                            robot.keyToggle('alt', 'down');
                            robot.keyTap('f4');
                            robot.keyToggle('alt', 'up');


                            setTimeout(() => {
                                robot.keyTap("enter");

                                setTimeout(() => {
                                    robot.keyTap("enter");

                                    setTimeout(() => {

                                        robot.keyTap("left")

                                        setTimeout(()=>{

                                            robot.keyTap("enter");
                                            //firebaseHelper.admin.database().ref("listener_documents").child("document").remove()
                                            const workbook = new Excel.Workbook()
                                            workbook.xlsx.readFile(`${__dirname}/${path}`).then(() => {

                                                const worksheet = workbook.worksheets[0]

                                                let result = [
                                                    {
                                                        score: `${worksheet.getCell('C425').value.result}`,
                                                        scoreText: `${worksheet.getCell('C427').value.result}`,
                                                        cell: `D424`
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C433').value.result}`,
                                                        scoreText: `${worksheet.getCell('C435').value.result}`,
                                                        cell: "D432"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C435').value.result}`,
                                                        scoreText: `${worksheet.getCell('C442').value.result}`,
                                                        cell: "D439"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C447').value.result}`,
                                                        scoreText: `${worksheet.getCell('C449').value.result}`,
                                                        cell: "D446"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C454').value.result}`,
                                                        scoreText: `${worksheet.getCell('C456').value.result}`,
                                                        cell: "D453"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C461').value.result}`,
                                                        scoreText: `${worksheet.getCell('C463').value.result}`,
                                                        cell: "D460"
                                                    }, {
                                                        score: `${worksheet.getCell('C468').value.result}`,
                                                        scoreText: `${worksheet.getCell('C470').value.result}`,
                                                        cell: "D467"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C475').value.result}`,
                                                        scoreText: `${worksheet.getCell('C477').value.result}`,
                                                        cell: "D474"
                                                    }, {
                                                        score: `${worksheet.getCell('C482').value.result}`,
                                                        scoreText: `${worksheet.getCell('C484').value.result}`,
                                                        cell: "D481"
                                                    },
                                                    {
                                                        score: `${worksheet.getCell('C489').value.result}`,
                                                        scoreText: `${worksheet.getCell('C491').value.result}`,
                                                        cell: "D488"
                                                    }, {
                                                        score: `${worksheet.getCell('C496').value.result ? worksheet.getCell('C496').value.result : ''}`,
                                                        scoreText: `${worksheet.getCell('C498').value.result ? worksheet.getCell('C498').value.result : ''}`,
                                                        cell: "D495"
                                                    }
                                                ]

                                                firebaseHelper.admin.database().ref("results").child(data.id).set(JSON.stringify({results: result}))
                                            })


                                        }, 1000)




                                    }, 1000)

                                }, 1000)

                            }, 1000)

                        }, 1000)

                    }, 1000)


                }, 1000)
                robot.moveMouse(20, 2);
            }, 3000)

        })


    }

})


const download = (uri, filename, callback) => {
    const file = fs.createWriteStream(filename);
    http.get(uri, (response) => {
        response.pipe(file);
        callback()
    });
};




