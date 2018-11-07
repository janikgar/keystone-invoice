//UI Functions
//
//

const TIMER = "timer";

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Generate Invoice", "onInvoice")
    .addToUi();
}

function onInstall() {
  onOpen();
  SpreadsheetApp.getUi().alert(
    "Keystone Invoice",
    `The Keystone Invoice addon has been installed. To run the script functions,` +
    `go to the Add-ons menu, and trigger the functions from the Keystone Invoice submenu.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function modalPop(url, filename) {
  var tempHtml = HtmlService.createTemplateFromFile('modal');
  tempHtml["url"] = url;
  tempHtml["filename"] = filename;
  SpreadsheetApp.getUi().showModalDialog(tempHtml.evaluate(), "Keystone Addon");
}

function onInvoice() {
  console.time(TIMER);
  toast("Generating Invoice...");
  var pairs = adminSubs();
  var date = toOurDateString(new Date());
  pairs.push(['<<Date>>', date])
  var invoiceTemplate = DriveApp.getFileById('1ky_9rtYbMH-hkktyk38Pv4Ruwvjw5C0FlyA74l5Fjfc');
  var folder = setFileInfo();
  var newDocDrive = invoiceTemplate.makeCopy(folder);
  var newDoc = SpreadsheetApp.openById(newDocDrive.getId());
  var sheet = newDoc.getSheetByName('Invoice');
  var sheet_values = sheet.getDataRange().getValues();
  console.log(sheet_values)
  var invoiceNumber;
  for (var i = 0; i < sheet_values.length; i++) {
    console.log(sheet_values[i]);
    for (var j = 0; j < pairs.length; j++) {
      var new_values = sheet_values[i].map(
        function (old_value) {
          return old_value.toString().replace(pairs[j][0], pairs[j][1])
        }
      );
      if (pairs[j][0] == '<<Keystone Number>>') {
        invoiceNumber = pairs[j][1];
      }
      var policyClaimType, policy, claim, insured;
      switch(pairs[0]) {
        case "<<Reference Claim/Policy>>":
          policyClaimType = pairs[1];
        case "<<Policy Number>>":
          policy = pairs[1];
        case "<<Claim Number>>":
          claim = pairs[1];
        case "<<Insured>>":
          insured = pairs[1];
      }
      sheet_values[i] = new_values
    }
  }
  console.log(sheet_values);
  sheet.getDataRange().setValues(sheet_values);
  newDocDrive.setName('')
  var nameSuffix = getNameSuffix(policyClaimType, policy, claim, insured);
  newDocDrive.setName(`Invoice ${invoiceNumber} ${nameSuffix}`);
  finishAndAlert(newDocDrive);
}

function adminSubs() {
  console.log(`document: ${SpreadsheetApp.getActive().getId()}, user: ${Session.getActiveUser()}`)
  var sheet = SpreadsheetApp.getActive().getSheetByName('Admin Fields');
  var varLength = sheet.getDataRange().getNumRows();
  var pairRange = sheet.getRange(2, 1, varLength, 2).getValues();
  var d = new Date();
  var pairs = [];
  pairs.push(["<<Report Date>>", toOurDateString(d)]);
  for (let pair of pairRange) {
    if (pair[0] === "" || pair[1] === "") {
      if (pair[0] === "Property Address Line 2") {
        if (pair[1] != "") {
          pairs.push([`<<${pair[0].toString().trim()}>>`, `\n${pair[1]}`]);
        } else {
          pairs.push([`<<${pair[0].toString().trim()}>>`, " "]);
        }
      }
    } else if (pair[0] instanceof Date) {
      pairs.push([`<<${pair[1]}>>`, toOurDateString(pair[0])]);
    } else {
      pairs.push([`<<${pair[0].toString().trim()}>>`, pair[1]]);
    }
  }
  return pairs
}

function getNameSuffix(policyClaimType, policy, claim, insured) {
  var nameSuffix;
  switch (true) {
    case policyClaimType === "Policy":
      nameSuffix = `${policyClaimType} ${policy} (${insured})`;
      break;
    case policyClaimType === "Claim":
      nameSuffix = `${policyClaimType} ${claim} ${insured}`;
      break;
    default:
      nameSuffix = insured;
  }
  return nameSuffix
}


function finishAndAlert(newDoc) {
  Utilities.sleep(500);
  modalPop(newDoc.getUrl(), newDoc.getName());
  console.timeEnd(TIMER);
  return 0;
}

function setFileInfo() {
  var fileId = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
  if (fileId.getParents().hasNext()) {
    var folder = fileId.getParents().next();
    return folder
  } else {
    SpreadsheetApp.getUi().alert("Error: It appears that you do not have read-write access to the folder containing this Spreadsheet. Contact your administrator to make sure you can read and write in this folder on Drive. If this error message keeps occuring, contact the plugin author at janikgar@gmail.com.", SpreadsheetApp.getUi().ButtonSet.OK);
    return "1"
  }
}

// Utility functions
//
//
function toast(e) {
  SpreadsheetApp.getActive().toast(e, "Keystone Status", 2);
}

function toOurDateString(d) {
  var monthList = {
    0: "January",
    1: "February",
    2: "March",
    3: "April",
    4: "May",
    5: "June",
    6: "July",
    7: "August",
    8: "September",
    9: "October",
    10: "November",
    11: "December"
  }
  var month = monthList[d.getMonth()];
  var day = d.getDate();
  var year = d.getFullYear();
  var returnDate = `${month} ${day}, ${year}`;
  return returnDate
}