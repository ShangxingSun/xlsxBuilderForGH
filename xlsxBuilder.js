var xlsx = require('node-xlsx').default;
const fs = require('fs');

console.log(process.argv[2]);
const workSheetsFromFile = xlsx.parse("./"+process.argv[2]);

var sheetOutput = [];
var sheetToProcess = workSheetsFromFile[1];
console.log(sheetToProcess.name);
console.log(sheetToProcess.data);
var Header = ['Ticket Number','Ticket Link']
sheetOutput.push(Header);
sheetToProcess.data.forEach((ele)=>{
    var ticketNum = ele[0];
    var ticketLink = {v:ticketNum,t:'s',l:{Target:"https://gethired.atlassian.net/browse/"+ticketNum}};
    var line = [ticketNum,ticketLink];
    sheetOutput.push(line);
});
console.log(sheetOutput);
var buffer = xlsx.build([{name: "Weekly report", data: sheetOutput}]);
fs.writeFileSync("weeklyReport.xlsx", buffer, 'binary');

