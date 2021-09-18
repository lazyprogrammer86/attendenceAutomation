//-----------------------------------------------begining----------------------------------------//

const min_duration = 45;
const dateTime = '2021-08-24T08:30:00'
// Requiring the module
const xlsx = require('xlsx')

//------------------------------------------register.xlsx---------------------------------------//

// Reading test file registration
const rwb = xlsx.readFile('./register.xlsx')

// destructuring the array to take only the first element of the array
const [rsheets, ] = rwb.SheetNames

//converting sheet data into easily accessible json format
const temp_r = xlsx.utils.sheet_to_json(rwb.Sheets[rsheets])

//array of  objects
//console.log(temp_r)  


//------------------------------------------test.xlsx---------------------------------------//

//updating the register file using the data available in the test
const twb = xlsx.readFile('./test.xlsx')

//destructuring array
const [tsheets, ] = twb.SheetNames

//converting test sheet datas to json data
const temp_t = xlsx.utils.sheet_to_json(twb.Sheets[tsheets])

//array of test objects
//console.log(temp_t)

//------------------------------------updating the registration file --------------------------------//

//filtering out the email form the documents which do not meet the min duration criteria
const newTestData = (temp_t.filter(data => data.Duration >= min_duration)).map(email => email.Email_ID)

temp_r.forEach(regData => {
   var updateAttendence = newTestData.find(email_ID => regData.Email_ID === email_ID) //finding email in register based on available email in the newTestData array

   if (updateAttendence) {
      regData[dateTime] = 'present' // marking present in registration if email ID is present in the array
   }else{
      regData[dateTime]='absent' //marking absent if the email ID is not present in the array
   }
})

twb.Sheets[tsheets] = xlsx.utils.json_to_sheet(temp_r)

xlsx.writeFile(twb,'./register.xlsx')

//-----------------------------------------END--------------------------------------------------------//