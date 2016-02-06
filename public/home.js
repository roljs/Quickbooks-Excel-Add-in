intuit.ipp.anywhere.setup({
    grantUrl: 'https://agave-node.azurewebsites.net/requestToken', datasources: {
        quickbooks: true,  // set to false if NOT using Quickbooks API
        payments: false    // set to true if using Payments API
    }
});

Office.initialize = function (reason) {
    $(document).ready(function () {

        overrideWinOpen();

        $('#btnGetPurchases').click(getPurchases);
        $('#btnGetAccounts').click(getAccounts);
        $('#btnCreateReport').click(createReport);
        $('#btnSignOut').click(signOut);
        $('#btnSendToCollabDb').click(sendToCollabDb);
        $('#btnPostToCollabDb').click(postToCollabDb);

        init();
    });
};


var _dlg;
function overrideWinOpen() {

    window.open = function (URL, name, specs, replace) {
        Office.context.ui.displayDialogAsync(URL,
            { height: 40, width: 40, requireHTTPS: true },
            function (result) {
                _dlg = result.value;
                _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
            });
    }

}

function processMessage(arg) {
    if (arg.message == "success") {
        _dlg.close();
        $("#welcomePanel").hide();
        $("#actionsPanel").show();
    }
}

function init() {

    $.get("/getToken", function (data, status) {
        if (data.oauth_token_secret) {
            $("#welcomePanel").hide();
            $("#actionsPanel").fadeIn("slow");
        }
        else {
            $("#welcomePanel").fadeIn("slow");
            $("#actionsPanel").hide();
        }
    });
}

function signOut() {
    $.get("/clearToken", function (data, status) {
        init();
    });
}

function getPurchases() {
    $.get("/getPurchases", function (data, status) {
        createPurchasesTable(data.QueryResponse.Purchase);
    });

}

function createPurchasesTable(purchases) {
    Excel.run(function (ctx) {

        var sheet = ctx.workbook.worksheets.add("Expenses");
        sheet.activate();
        // Queue a command to add a new table
        var table = ctx.workbook.tables.add('Expenses!A2:E2', true);
        table.name = "Purchases";

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = [["Date", "Type", "Payee", "Category", "Amount"]];

        // Create a proxy object for the table rows
        var tableRows = table.rows;

        $.each(purchases, function (i, item) {
            var date = item.TxnDate;
            var type = item.PaymentType;
            var payee = "";
            if (item.EntityRef)
                payee = item.EntityRef.name;
            var cat = "";
            if (item.Line.length > 0) {
                switch (item.Line[0].DetailType) {
                    case "AccountBasedExpenseLineDetail":
                        cat = item.Line[0].AccountBasedExpenseLineDetail.AccountRef.name;
                        break;
                    case "ItemBasedExpenseLineDetail":
                        cat = item.Line[0].ItemBasedExpenseLineDetail.ItemRef.name;
                        break;

                }
            }
            var amount = item.TotalAmt;

            var r = tableRows.add(null, [[date, type, payee, cat, amount]]);
            r.getRange().numberFormat = [[null, null, null, null, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)']];
            r.getRange().format.autofitColumns();

            addTitle(sheet, "A1:E1", "A1", "Expense");

        });



        return ctx.sync()

    }).catch(function (error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}

function addTitle(sheet, range, start, titleText) {

    var title = sheet.getRange(range);
    title.format.fill.color = "336699";
    title.format.font.color = "white";
    title.format.font.size = 24;
    title = sheet.getRange(start);
    title.values = titleText;

}

function getAccounts() {
    $.get("/getAccounts", function (data, status) {
        createAccountsTable(data.QueryResponse.Account);
    });

}

function createAccountsTable(accounts) {

    Excel.run(function (ctx) {

        var sheet = ctx.workbook.worksheets.add("Accounts");
        sheet.activate();
        
        // Queue a command to add a new table
        var table = ctx.workbook.tables.add('Accounts!A2:E2', true);
        table.name = "Accounts";

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = [["Name", "Currency", "Type", "Class", "Balance"]];

        // Create a proxy object for the table rows
        var tableRows = table.rows;

        $.each(accounts, function (i, item) {
            var r = tableRows.add(null, [[item.Name, item.CurrencyRef.value, item.AccountType, item.Classification, item.CurrentBalance]]);
            r.getRange().numberFormat = [[null, null, null, null, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)']];
            r.getRange().format.autofitColumns();

            addTitle(sheet, "A1:E1", "A1", "Accounts");

        });


        return ctx.sync()

    }).catch(function (error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}


function createReport() {
    Excel.run(function (ctx) {
        var sheet = ctx.workbook.worksheets.add("Spending Report");
        var address = "A2:B5";

        sheet.activate();

        var sumRange = sheet.getRange(address);
        sumRange.values = [['Type', 'Total'],
            ['Credit Card', '=SUMIF( Expenses!B:B, "CreditCard", Expenses!E:E )'],
            ['Check', '=SUMIF( Expenses!B:B, "Check", Expenses!E:E )'],
            ['Cash', '=SUMIF( Expenses!B:B, "Cash",Expenses!E:E )']];
        var currencyFormat = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
        sumRange.numberFormat = [[null, null], [null, currencyFormat], [null, currencyFormat], [null, currencyFormat]];
        sumRange.format.autofitColumns();

        ctx.workbook.tables.add(address, true);

        var chartRange = sheet.getRange(address);
        var chart = ctx.workbook.worksheets.getActiveWorksheet().charts.add("Doughnut", chartRange);
        chart.title = "Spending by Type";
        addTitle(sheet, "A1:E1", "A1", "Spending Report");

        return ctx.sync();

    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

var accessToken = "";
function sendToCollabDb() {
    if (accessToken == "") {
        Office.context.ui.displayDialogAsync("https://agave-node.azurewebsites.net/aadAuth.html",
            { height: 40, width: 40, requireHTTPS: true },
            function (result) {
                _dlg = result.value;
                _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (arg) {
                    var message = JSON.parse(arg.message); 
                    if (message.status == "success") {
                        _dlg.close();
                        accessToken = message.accessToken;
                        $('#collabDbPanel').show("slow");

                    }
                });
            });

    }else{
        $('#collabDbPanel').toggle();
    }
    
}