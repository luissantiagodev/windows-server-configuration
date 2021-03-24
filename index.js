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

        let path = `${data.id}.xls`
        let command = `${EXCEL_EXE} "${__dirname}/${path}"`

        console.log(path)
        console.log(command)
        console.log(__dirname)

        download(data.url, path, () => {
            cmd.run(command, (err, data, stderr) => {
                    firebaseHelper.admin.database().ref("listener_documents").child("document").remove()
                }
            );

            setTimeout(() => {
                robot.mouseClick();
                robot.keyTap("left");
                robot.keyTap("enter");
            }, 2000)
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