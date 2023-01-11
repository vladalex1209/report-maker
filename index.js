const fs = require('fs');
const prompt = require('prompt-sync')({ sigint: true });
const { toXlsxFile } = require('./xlsx.js');
const { fromHubstaff } = require("./selenium.js");
const { promptUser } = require('./prompts.js');

(async function getReport() {
  try {

    const envFilepath = './.env'

    if (!fs.existsSync(envFilepath)) {
      let email = await promptUser('Please enter your Hubstaff email: ');
      let password = await promptUser('Please enter your Hubstaff password: ', true);
      let user = await promptUser('Please enter your computer username: ');
      fs.appendFileSync('.env', `EMAIL=${email}\nPASSWORD=${password}\nUSER=${user}`);
      console.log('Your credentials have been saved in .env file. Please run the script again.')
      process.exit();
    }

    currentDate = new Date();

    currentDate = currentDate.toISOString().split('T')[0];

    let reportObj = [];

    await fromHubstaff(reportObj, currentDate)
    console.log(reportObj)
    toXlsxFile(reportObj, currentDate);

    console.log(`Report was created succesfully in the reports folder with name : report - ${currentDate}.txt`);

  }
  catch (err) {
    console.error(err);
    if (driver) driver.close();
  }
})();

process.on('SIGINT', function () {
  console.log("Caught interrupt signal");
  process.exit();
});





