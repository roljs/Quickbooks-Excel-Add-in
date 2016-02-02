var _open = window.open;
var _timer;
//window.open = function (URL,name,specs,replace) { window.location.href = URL; }

Office.initialize = function (reason) {
    $(document).ready(function () {

        overrideWinOpen();
        _timer = setInterval(checkToken, 1000);
        $('#btnGetAccounts').click(getAccounts);
        $('#btnGetPurchases').click(getPurchases);


    });
};

function checkToken() {

    $.get("/token", function (data, status) {
        if (data.oauth_token_secret) {
            clearInterval(_timer);
            $(".intuitPlatformConnectButton").hide();
        }
    });
}

function overrideWinOpen() {

    window.open = function (URL, name, specs, replace) {
        Office.context.ui.displayDialogAsync(URL,
            { height: 40, width: 40, requireHTTPS: true },
            function () { });
    }

}

function getAccounts() {
    $.get("/getAccounts", function (data, status) {
        createAccountsTable(data.QueryResponse.Account);
    });

}

function getPurchases() {
    $.get("/getPurchases", function (data, status) {
        createPurchasesTable(data.QueryResponse.Purchase);
    });

}

function createAccountsTable(accounts) {


}

function createPurchasesTable(purchases) {
    Excel.run(function (ctx) {

        var sheet = ctx.workbook.worksheets.add("Expenses");
        
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
            
            //var usedRange = table.getRange();
            //usedRange./*getEntireColumn().*/format.autofitColumns();
            //usedRange.getEntireRow().format.autofitRows();

            
            var title = sheet.getRange("A1:E1");
            title.format.fill.color = "336699";
            title.format.font.color = "white";
            title.format.font.size = 24;
            title = sheet.getRange("A1");
            title.values = "Expenses";
            

        });
        return ctx.sync().then(function (ctx) {



        })
    }).catch(function (error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        //app.showNotification("Error: " + error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}