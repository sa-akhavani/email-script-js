var XLSX = require('xlsx');
var fs = require('fs');
var nodemailer = require('nodemailer');

var nameCol = "A";
var emailCol = "C";

// Notice That this code does not support more than 15 rows because of limited number of sessions.
// I am going to make the sendMail function to syncronus in order to solve this problem.
var numOfRows = 10;
var fileName = "input.xlsx";

var workbook = XLSX.readFile(fileName);
var sheet1 = workbook.SheetNames[0];
var worksheet = workbook.Sheets[sheet1];

var person = [];

var htmlEmail = '';
var htmlEmailPrefix = fs.readFileSync("htmlPrefix.html");
var htmlEmailSuffix = fs.readFileSync("htmlSuffix.html");


for (var i = 2; i <= numOfRows; i++) {
	var name = worksheet[nameCol + String(i)];
	var email = worksheet[emailCol + String(i)];
	var newPerson = {
		name: name.v,
		email: email.v
	};
	person.push(newPerson);
}

var transporter = createTransporter();

for (var i = 0; i < person.length; i++) {
	transporter.sendMail({
		from: "Seyyed Ali Akhavani <***@gmail.com>",
		to: person[i].email,
		subject: "Happy New Year!",
		html: htmlEmailPrefix + person[i].name + htmlEmailSuffix
	}, function (err, info) {
		if (err)
			console.log("err: " + err + "" + person[i].name);
		else
			console.log("Message Sent: " + info.response + " | " + person[i].name);
	});
}


function createTransporter() {
	var transporter = nodemailer.createTransport({
		service: 'Gmail',
		auth: {
			user: '***@gmail.com',
			pass: "***"
		}
	});
	return transporter;
}