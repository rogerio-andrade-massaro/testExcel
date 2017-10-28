(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    var root;

    //http://jsfiddle.net/hybrid13i/JXrwM/

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");


                return;
            }



            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(loadSampleData);

        });
    };

    function JSONToArray(JSONData) {
        var arr = [];

        arr.push(_.map(JSONData[0], function (num, key) { return key }));

        for (var i = 0; i < JSONData.length; i++) {
            arr.push(_.map(JSONData[i], function (num, key) { return num }));
        }

        //https://github.com/OfficeDev/Excel-Add-in-JS-ExternalDataGitHub/blob/master/Code%20Editor%20Proj/Home.js

        return arr;
    }

    function loadSampleData() {

        $.ajax({
            url: $("#txtURL").val(),
            method: 'GET'
        }).then(function (values) {
            values = JSONToArray(values);

            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {
                // Create a proxy object for the active sheet
                // https://dev.office.com/docs/add-ins/excel/excel-add-ins-tables?product=excel

                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                var rangeHeader = sheet.getCell(0, 0).getBoundingRect(sheet.getCell(0, values[0].length - 1));
                var newrange = sheet.getCell(0, 0).getBoundingRect(sheet.getCell(values.length - 1, values[0].length - 1));
                //newrange.format.fill.color = "gray";
                // Queue a command to write the sample data to the worksheet
                newrange.clear;
                newrange.values = values;
                newrange.format.autofitColumns;
                newrange.format.autofitRows;
                rangeHeader.format.fill.color = "blue";
                rangeHeader.format.font.bold = true;
                rangeHeader.format.font.color = "white";
                sheet.getUsedRange.autofitColumns;
                
                var expensesTable = sheet.tables.add('A1:C101', true); //convert data to table
                expensesTable.name = "ExpensesTable";
                expensesTable.autofitColumns;
                expensesTable.autofitRows;

                sheet.getUsedRange.autofitColumns;

                // Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync().then(showNotification('Data loaded'));
            })
                .catch(errorHandler);
        });


    }

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);


    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
