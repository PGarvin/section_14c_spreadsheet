const url1 = `https://www.dol.gov/sites/dolgov/files/WHD/xml/CertificatesListing.xlsx`;

const time_stamp = new Date() - new Date("June 1, 2023");

const { ApiKeyManager } = require('@esri/arcgis-rest-request');
const { geocode } = require('@esri/arcgis-rest-geocoding');

const apiKey = "API KEY GOES HERE";
const authentication = ApiKeyManager.fromKey(apiKey);

const reader = require('xlsx')
const https = require('https')
const fs = require('fs');

https.get(url1, resp => resp.pipe(fs.createWriteStream('./raw_data.xlsx')));

const checkTime = 1750;
const messageFile = "./raw_data.xlsx";
const timerId = setInterval(() => {
  const isExists = fs.existsSync(messageFile, 'utf8');
  if(isExists) {
    clearInterval(timerId);
    readExcelFile(messageFile);
  }
}, checkTime);

const updatedData = [];

let numberOfWorkersPaidSubminimumWages = 0;
let numberOfLatLong = 0;

function readExcelFile(fileName) {

  const file = reader.readFile(fileName)

  const data = [];
  const info = [];

  const sheets = file.SheetNames;

     const temp = reader.utils.sheet_to_json(
          file.Sheets[file.SheetNames[0]])
     temp.forEach((res) => {
        data.push(res);
        if (res['Number of Workers Paid Subminimum Wages'] !== undefined) {
        numberOfWorkersPaidSubminimumWages+= Number(res['Number of Workers Paid Subminimum Wages']);
        }
     });

    var states;


    let keys = Object.keys(data[0]);
    const keyNames = [];
    keys.forEach((keyName) => {
      //console.log(keyname);
      if (keyName.indexOf("SCA") < 0 && keyName.indexOf("PCA") < 0 && keyName.indexOf("Renewal") < 0 && keyName.indexOf("Address") < 0 && keyName.indexOf("Date") < 0 && keyName.indexOf("Number") < 0 && keyName.indexOf("Zipcode") < 0 && keyName.indexOf("City") < 0) {
        console.log(keyName);
        keyNames.push(keyName);
        info[keyName] = [];
      }
    });

    data.forEach((datum, index) => {

      geocode({
  address: datum.Address,
  postal: datum.Zipcode,
  countryCode: "USA",
  authentication,
}).then((response) => {
    response.candidates.forEach((res) => {
      //console.log(68, res);
      data[index].longitude = res.location.x;
      data[index].latitude = res.location.y;
      updatedData.push(data[index])
    })
});


      keyNames.forEach(keyName => {
        let currentKey = datum[keyName];
        if (isNaN(currentKey) === false) {
          //currentKey = JSON.stringify(currentKey);
        }
        let doesKeyExist = 0;
        for (var j = 0; j < info[keyName].length; j++) {
          if (info[keyName][j].name === currentKey) {
            doesKeyExist = 1;
            info[keyName][j].value++;
            if (keyName === "State" && datum['Number of Workers Paid Subminimum Wages'] !== undefined) {
              info[keyName][j]['Number of Workers Paid Subminimum Wages'] = info[keyName][j]['Number of Workers Paid Subminimum Wages'] + Number(datum['Number of Workers Paid Subminimum Wages']);
            }
          }
        }

        if (doesKeyExist < 1) {
          if (keyName !== "State") {
            info[keyName].push({"name":currentKey,"value":1});
          } else if (keyName === "State" && datum['Number of Workers Paid Subminimum Wages'] !== undefined) {
            info[keyName].push({"name":currentKey,"value":1,'Number of Workers Paid Subminimum Wages':Number(datum['Number of Workers Paid Subminimum Wages'])});
          }
          doesKeyExist = 1;
        }
      });
    });


    const  checkTimerId = setInterval(() => {
      console.log(updatedData.length, data.length, updatedData.length/data.length);
      if(updatedData.length >= data.length) {
        const wb = reader.utils.book_new()

        keyNames.forEach(keyName => {

        info[keyName].sort((a, b) => (a.value < b.value) ? 1 : -1)

        reader.utils.book_append_sheet(wb, reader.utils.json_to_sheet(info[keyName]), keyName.split("/").join("").split(" ").join("_"))

        });
        reader.utils.book_append_sheet(wb, reader.utils.json_to_sheet(data, "overall_data"));

        reader.writeFile(wb, 'updated_data_with_lat_lon.xlsx');
        clearInterval(checkTimerId);
      }
    }, 3050)



}
