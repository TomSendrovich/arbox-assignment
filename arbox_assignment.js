var fs = require ('fs');
var NodeXls = require('node-xls');

//Need to configure those 5 params for the script to work for many clubs and computers
var fullInputClubPath = 'C:\\Users\\Tsand\\OneDrive\\Desktop\\jimalaya.xlsx';
var fullInputAr_DBPath = 'C:\\Users\\Tsand\\OneDrive\\Desktop\\ar_db.xlsx';
var clubId = "2400";
var fullUsersOutputPath = "C:\\Users\\Tsand\\OneDrive\\Desktop\\userOutput.xlsx"
var fullMembershipOutputPath = "C:\\Users\\Tsand\\OneDrive\\Desktop\\membershipOutput.xlsx"

import parseXlsx from 'excel';

//Checks for duplicate emails, if true exit program whitout creating the files,
// else continue regularly
function isDuplicate(email, emailArr){
  var isDuplicateEmail = emailArr.includes(email);
  if (isDuplicateEmail){
    console.log("Found duplicate email, exit program");
    process.exit(1);
  }
  return;
}

//Converts excel time to regular date.
function excelTimeToDate(excelTime){
  var date = new Date(Math.round((excelTime - (25567 + 2)) * 86400 * 1000));
  return date.toLocaleDateString();
}

var user_id;
//finds the next next user_id to put in the DB:
try{
  parseXlsx(fullInputAr_DBPath).then((data) => {
    user_id = data.length;
  });
}
catch(err){
  console.log("Can't find ar_db.xlsx file. exit program");
  process.exit(1);
}

//Iterating through the club's file, extracting the data,
// and creating file to import to DB
try{
  parseXlsx(fullInputClubPath).then((data) => {
    // data is an array of arrays of jimalaya.xlsx
    var usersData = []
    var membershipData = []
    var index = 0;
    var emailArr = []
    if (user_id == undefined){
      user_id = 4;
    }
    //Loops through the data to convert it to a .json file
    for (let i=1;i< data.length; i++){
      var firstName = data[i][0];
      var lastName = data[i][1];
      var email = data[i][2];
      isDuplicate(email,emailArr);
      emailArr.push(email);
      var phone = data[i][3];
      var startDate = data[i][4];
      startDate = excelTimeToDate(startDate);
      var endDate = data[i][5];
      endDate = excelTimeToDate(endDate);
      var membership = data[i][6];
      usersData[index] = {first_name:firstName, last_name:lastName, phone:phone,
         email:email, joined_at:startDate, club_id:clubId};
      membershipData[index] = {user_id:user_id, start_date:startDate, end_date:endDate,
         membership_name:membership};
      index++;
      user_id++;
    }
    var tool = new NodeXls();

    // Convert the .json file to excel with the correlated columns
    var userExcel = tool.json2xls(usersData, {order:["first_name", "last_name",
     "phone", "email", "joined_at", "club_id"]});
    fs.writeFileSync(fullUsersOutputPath,userExcel, 'binary');

    var membershipExcel = tool.json2xls(membershipData, {order:["user_id", "start_date",
     "end_date", "membership_name"]});
    fs.writeFileSync(fullMembershipOutputPath,membershipExcel, 'binary');
  });
}
catch(err){
  console.log("Error while creating file for DB.")
  throw err;
}
