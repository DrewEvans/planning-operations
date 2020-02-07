function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Merchandising Tools')
        .addItem('Freshdesk Submission Review', 'showurl')
        .addSubMenu(ui.createMenu('More Functions')
            .addItem('Auto Plan My Week', 'rickRolled')
            .addItem('Reset Merchandising Calendar', 'refreshData'))
        .addToUi();
}

function showurl() {
    var ss = SpreadsheetApp.getActiveSheet();
    var sheet = SpreadsheetApp.getActive().getSheetByName("Daily Freshdesk Checks");
    var chart = sheet.getCharts()[0];
    var chartImage = chart.getAs("image/jpeg");


    var today = new Date();
    var time =
        today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();

    var template = HtmlService.createTemplateFromFile("reportTemplate");

    var htmlOutput = HtmlService.createHtmlOutputFromFile("reportTemplate")
        .setWidth(700) //optional
        .setHeight(600); //optional
    htmlOutput.append(
        "<p align='left'><img src='data:image/jpg;base64," +
        Utilities.base64Encode(chartImage.getBytes()) +
        "'/>"
    );
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Daily Notification");
    Logger.log("chart");
}

function rickRolled() {
    var ss = SpreadsheetApp.getActiveSheet();
    Browser.msgBox('CONGRATULATIONS', 'Your week looks the exact same before you pressed this button', Browser.Buttons.OK);
}

function refreshData() {
    var countryTabs = [
        'UK',
        'FR',
        'DE',
        'IT',
        'ES',
        'NL',
        'BE',
        'IE',
        'PL',
        'AE'
    ];
    var rangesToClear = [
        'E6:GD9',
        'E11:GD20',
        'E23:GD26',
        'E28:GD37',
        'E40:GD43',
        'E45:GD54',
        'E57:GD60',
        'E62:GD71',
        'E74:GD77',
        'E79:GD88',
        'E91:GD94',
        'E96:GD105',
        'E108:GD113',
        'E116:GD117',
        'E119:GD124',
        'E127:GD128',
        'E130:GD135',
        'E138:GD139',
        'E141:GD144',
        'E148:GD151',
        'E153:GD156',
        'E160:GD163',
        'E166:GD168',
        'E172:E175'
    ];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    countryTabs.forEach(function (sheetName) {
        sheetToClear = ss.getSheetByName(sheetName);

        //If any sheets/tabs are listed in countryTabs then perform the below 
        if (sheetToClear) {
            for (var i = 0; i < rangesToClear.length; i++) {
                //Clears content within cells listed in rangesToClear
                sheetToClear.getRange(rangesToClear[i]).clearContent();
                //Clears notes created within cells listed in RangesToClear
                sheetToClear.getRange(rangesToClear[i]).clearNote();
                //Clears any comments Created listed in rangesToClear *bug reported https://issuetracker.google.com/issues/36756650
                sheetToClear.getRange(rangesToClear[i]).clear({
                    commentsOnly: true
                });
                //Reverts Cells listed in RangesTo Clear to set cell colour white
                sheetToClear.getRange(rangesToClear[i]).setBackgroundColor('white');
                Logger.log(sheetToClear);
            }
        }
    })
}
