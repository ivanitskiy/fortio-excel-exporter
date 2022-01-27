'use strict';
const writeXlsxFile = require('write-excel-file/node')
const fs = require('fs');
const url = require('url');
const glob = require('glob');

const headerData = [
  { value: "URL" },
  { value: "NumThreads" },
  { value: "StartTime" },
  { value: "Labels" },
  { value: "RequestedQPS" },
  { value: "RequestedDuration" },
  { value: "ActualDuration" },
  { value: "ActualQPS" },
  { value: "SocketCount" },
  { value: "Count" },
  { value: "Min" },
  { value: "Max" },
  { value: "Sum" },
  { value: "Avg" },
  { value: "StdDev" },
  { value: "p50" },
  { value: "p75" },
  { value: "p90" },
  { value: "p99" },
  { value: "p999" },
  { value: "Code -1" },
  { value: "Code 0" },
  { value: "Code 200" },
  { value: "Code 204" },
  { value: "Code 401" },
  { value: "Code 403" },
  { value: "Code 404" },
  { value: "Code 502" },
  { value: "Code 503" },
  { value: "Code 511" },
  { value: "Filename" }
];


async function saveXls(d, s) {
  const outputStream = fs.createWriteStream('test-stream.xlsx');

  await new Promise((resolve, reject) => {
    writeXlsxFile(d, {
      sheets: s,
      headerStyle: {
        backgroundColor: '#eeeeee',
        fontWeight: 'bold',
        align: 'center'
      },
    }).then((stream) => {
      stream.pipe(outputStream)
      stream.on('end', function () {
        console.log('XLSX file stream ended')
      })
    })
    outputStream.on('close', function () {
      console.log('Output stream closed')
      resolve()
    })
  })
}

function readResultFromFile(filePath) {
  let rawData = fs.readFileSync(filePath);
  let runResult = JSON.parse(rawData);

  return {
    "URL": runResult.URL,
    "Labels": runResult.Labels,
    "StartTime": runResult.StartTime,
    "RequestedQPS": runResult.RequestedQPS,
    "RequestedDuration": runResult.RequestedDuration,
    "ActualQPS": runResult.ActualQPS,
    "ActualDuration": runResult.ActualDuration,
    "NumThreads": runResult.NumThreads,
    "Count": runResult.DurationHistogram.Count,
    "Min": runResult.DurationHistogram.Min,
    "Max": runResult.DurationHistogram.Max,
    "Sum": runResult.DurationHistogram.Sum,
    "Avg": runResult.DurationHistogram.Avg,
    "StdDev": runResult.StdDev,
    "p50": runResult.DurationHistogram.Percentiles[0].Value,
    "p75": runResult.DurationHistogram.Percentiles[1].Value,
    "p90": runResult.DurationHistogram.Percentiles[2].Value,
    "p99": runResult.DurationHistogram.Percentiles[3].Value,
    "p999": runResult.DurationHistogram.Percentiles[4].Value,
    "SocketCount": runResult.SocketCount,
    "RetCodes": runResult.RetCodes,
    "Filename": filePath
  }
}

function XLSXformat(d) {
  let rows = [];
  let sheets = [];

  for (const [hostname, value] of Object.entries(d)) {
    let data = [
      headerData,
    ]
    for (const element of value) {
      data.push([
        { value: element.URL },
        { value: element.NumThreads },
        { value: element.StartTime },
        { value: element.Labels },
        { value: element.RequestedQPS },
        { value: element.RequestedDuration },
        { value: element.ActualDuration },
        { value: element.ActualQPS },
        { value: element.SocketCount },
        { value: element.Count },
        { value: element.Min },
        { value: element.Max },
        { value: element.Sum },
        { value: element.Avg },
        { value: element.StdDev },
        { value: element.p50 },
        { value: element.p75 },
        { value: element.p90 },
        { value: element.p99 },
        { value: element.p999 },
        { value: element.RetCodes["-1"] || 0 },
        { value: element.RetCodes["0"] || 0 },
        { value: element.RetCodes["200"] || 0 },
        { value: element.RetCodes["204"] || 0 },
        { value: element.RetCodes["401"] || 0 },
        { value: element.RetCodes["403"] || 0 },
        { value: element.RetCodes["404"] || 0 },
        { value: element.RetCodes["502"] || 0 },
        { value: element.RetCodes["503"] || 0 },
        { value: element.RetCodes["511"] || 0 },
        { value: element.Filename }
      ]);
    }
    sheets.push(hostname)
    rows.push(data);
  }
  return {
    data: rows,
    sheets: sheets
  }

}

glob('results/**/*.json', (err, files) => {
  if (err) {
    console.log(err);
  } else {
    let data = {};
    for (const filename of files) {
      let r = readResultFromFile(filename);
      let currentUrl = new url.URL(r.URL);
      let hostname = currentUrl.hostname;
      if (data[hostname] === undefined) {
        data[hostname] = []
      }
      data[hostname].push(r)
    }
    let xls = XLSXformat(data);
    console.log(xls.sheets);
    console.log(JSON.stringify(xls.data));
    saveXls(xls.data, xls.sheets);
  }
});
