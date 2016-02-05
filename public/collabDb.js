var cols = [];
var colPropertyNames = [];
var colTypes = [];
var currentTableId;
var tableLink;
var data;
var selection;
var errorOccurred;
var checkboxChecked;

function postToCollabDb() {
    console.log("getDataFromSelection called");
    $('#spinner').show();

    checkboxChecked = $('#headers-box').prop('checked');
    clearExistingValues();
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,{"filterType":"onlyVisible"},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                //$('#display-data').text('The selected text is: ');
                //prettyPrint(result.value);
                selection = result.value;
                beginImport();
            } else {
                $('#display-data').text('Error: ' + result.error.message);
            }
        }
        );
}

function beginImport() {
    //parse selection into format
    var colSize = selection[0].length;
    var rowSize = selection.length;
    cols = [];
    if (checkboxChecked) {
        cols = selection[0];
    } else {
        var i;
        for (i = 0; i < colSize; i++) {
            cols.push("Column " + (i + 1));
        }
    }
    data = selection;

    if (!errorOccurred && accessToken != null) {
        $('#aux').text("Adding schema...");
        createTable();
        if (!errorOccurred && currentTableId != undefined) {
            addData();
            $('#spinner').hide();
            $('#tableLink').text("Here's your new table");
            $('#tableLink').attr('href', tableLink);
        } else {
            $('#aux').text("Error occurred while creating table schema error:" + errorOcurred + " tableId:" + currentTableId);
        }
    } else {
        errorOccurred = true;
        $('#aux').text("Token error!");
    }
}

function addData() {
    $('#aux').text("Adding data...");
    var count = 0;
    if (checkboxChecked) {
        count = 1;
    }
    $('#display-data').text("data count is: " + data.length);
    var batch = "";
    var batchId = "12345";

    for (count; count < data.length; count++) {
        //addRow(data[count]);
        batch += getBatchRow(data[count], currentTableId, batchId);

    }
    sendBatch(batch, batchId);
}

function sendBatch(batch, batchId) {
    $.ajax({
        type: "POST",
        async: false,
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'multipart/mixed; boundary=batch_' + batchId
        },
        url: "https://lists.office.com/v1.0/fda85edf-d0c7-4d53-b3b8-7d1f3955fa1c/users/f03dacd2-d66a-4f09-b61f-30263a610646/$batch",
        data: batch,
        success: function (data, status, jqXhr) {
            console.log("Post successful");
        },
        error: function (jqXhr, status, error) {
            console.log(status);
        }
    });

}

function createTable() {
    var name = "ExcelTable " + getDateTime();
    $.ajax({
        type: "POST",
        async: false,
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        },
        url: "https://lists.office.com/v1.0/me/tableDefinitions/",
        data: JSON.stringify({ "displayName": name, "defaultAppRestriction": "None" }),
        dataType: 'json',
        success: function (msg) {
            currentTableId = msg.id;
            if (currentTableId != undefined) {
                patchRestriction(); // this should be a stop-gap
                tableLink = msg.tableLink;
                formatTableLink();
                var count;
                for (count = 0; count < cols.length; count++) {
                    var num = count + 1;
                    var firstValue = data[0][count];
                    var lastValue = data[data.length - 1][count];
                    var type = inferType(firstValue, lastValue);
                    colTypes.push(type);
                    addColumn(cols[count], type);
                }
            }
        },
        error: function (msg) {
            errorOccurred = true;
            $('#aux').text("Table creation failed");
        }
    });
}

function inferType(firstVal, lastVal) {
    //tests first and last point - if agreement, assume column is of that type
    if (checkboxChecked) {
        var lastValType = typeof lastVal;
        var dateLikelyforLastVal = lastVal > 10000 && lastVal < 45000;
        if (typeof lastVal == "number") {
            var dateLikelyforLastVal = lastVal > 10000 && lastVal < 45000;
            if (dateLikelyforLastVal) {
                return "Edm.DateTimeOffset";
            } else {
                return "Edm.Double";
            }
        } else if (typeof lastVal == "boolean") {
            return "Edm.Boolean";
        }
        return "Edm.String";
    } else {
        var firstValType = typeof firstVal;
        var lastValType = typeof lastVal;
        if (firstValType == lastValType) {
            if (typeof firstVal == "number") {
                var dateLikelyforFirstVal = firstVal > 10000 && firstVal < 45000;
                var dateLikelyforLastVal = lastVal > 10000 && lastVal < 45000;
                if (dateLikelyforFirstVal && dateLikelyforLastVal) {
                    return "Edm.DateTimeOffset";
                } else {
                    return "Edm.Double";
                }
            } else if (typeof firstVal == "boolean") {
                return "Edm.Boolean";
            }
        }
        return "Edm.String";
    }
}
function addColumn(title, type) {
    $.ajax({
        type: "POST",
        async: false,
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        },
        dataType: 'json',
        url: "https://lists.office.com/v1.0/me/tableDefinitions/" + currentTableId + "/columnDefinitions",
        data: JSON.stringify({ "displayName": title, "type": type }),
        success: function (msg) {
            colPropertyNames.push(msg.propertyName);
        }, error: function (msg) {
            errorOccurred = true;
            $('#aux').text("Column creation failed");
        }
    });
}

function addRow(data) {
    var vals = {};
    var i;
    for (i = 0; i < cols.length; i++) {
        var currentColumnType = colTypes[i];
        if (currentColumnType == "Edm.Double" || currentColumnType == "Edm.Boolean") {
            vals[colPropertyNames[i]] = data[i];
        } else if (currentColumnType == "Edm.DateTimeOffset") {
            var daysSince1900 = data[i];
            var daysBetween1900and1970 = 25568; // days between january 1st, 1900 and january 1st, 1970 according to wolframalpha
            var days = daysSince1900 - daysBetween1900and1970;
            var dt = new Date(days * 24 * 60 * 60 * 1000);
            vals[colPropertyNames[i]] = dt.toISOString();
        } else { //string default
            vals[colPropertyNames[i]] = String(data[i]);
        }
    }
    $.ajax({
        type: "POST",
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        },
        dataType: 'json',
        url: "https://lists.office.com/v1.0/me/tables/" + currentTableId + "/rows",
        data: JSON.stringify(vals),
        success: function (msg) {
        }, async: false,
        error: function (msg) {
            errorOccurred = true;
            $('#aux').text("Row creation failed");
        }
    });
}

function getBatchRow(data, tableId, batchId) {
    var vals = {};
    var i;
    for (i = 0; i < cols.length; i++) {
        var currentColumnType = colTypes[i];
        if (currentColumnType == "Edm.Double" || currentColumnType == "Edm.Boolean") {
            vals[colPropertyNames[i]] = data[i];
        } else if (currentColumnType == "Edm.DateTimeOffset") {
            var daysSince1900 = data[i];
            var daysBetween1900and1970 = 25568; // days between january 1st, 1900 and january 1st, 1970 according to wolframalpha
            var days = daysSince1900 - daysBetween1900and1970;
            var dt = new Date(days * 24 * 60 * 60 * 1000);
            vals[colPropertyNames[i]] = dt.toISOString();
        } else { //string default
            vals[colPropertyNames[i]] = String(data[i]);
        }
    }

    var batch = "";
    batch += "--batch_" + batchId + "\n";
    batch += "Content-Type: application/http\n";
    batch += "Content-Transfer-Encoding: binary\n\n";
    batch += "POST tables/" + tableId + "/rows HTTP/1.1\n";
    batch += "Accept: application/json\n";
    batch += "Content-Type: application/json\n\n";
    batch += JSON.stringify(vals) + "\n\n";
    
    return batch;
}

//for the time being:
function patchRestriction() {
    $.ajax({
        type: "PATCH",
        async: true, //this can be sent whenever
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        },
        url: "https://lists.office.com/v1.0/me/tableDefinitions/" + currentTableId,
        data: JSON.stringify({ "defaultAppRestriction": "None" }),
        dataType: 'json',
    });
}

//styling 2d array
function prettyPrint(array) {
    var i;
    var j;
    var temp = "";
    for (i = 0; i < array.length; i++) {
        for (j = 0; j < array[0].length; j++) {
            temp = temp + " " + array[i][j]
        }
        if (i + 1 != array.length) {
            temp = temp + "  /  ";
        }
    }
    $("#display-data").text(temp);
}

//sql styple time
function getDateTime() {
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth() + 1;
    var day = now.getDate();
    var hour = now.getHours();
    var minute = now.getMinutes();
    var second = now.getSeconds();
    if (month.toString().length == 1) {
        var month = '0' + month;
    }
    if (day.toString().length == 1) {
        var day = '0' + day;
    }
    if (hour.toString().length == 1) {
        var hour = '0' + hour;
    }
    if (minute.toString().length == 1) {
        var minute = '0' + minute;
    }
    if (second.toString().length == 1) {
        var second = '0' + second;
    }
    return dateTime = year + '/' + month + '/' + day + ' ' + hour + ':' + minute + ':' + second;
}

//major stop gap - table link in response should be of correct format.
function formatTableLink() {
    var parts = tableLink.split("https://lists.office.com/v1.0/");
    var tail = parts[1];
    var tailParts = tail.split("/");
    var tenantId = tailParts[0];
    var userId = tailParts[2];
    var tableId = tailParts[4];
    tableLink = "https://www.collabdb.com/?TenantID=" + tenantId + "&UserID=" + userId + "&TableID=" + tableId;
}

function clearExistingValues() {
    //deselect
    $("aux").text("working...");
    colPropertyNames = [];
    colTypes = [];
    cols = [];
    $('#tableLink').text("");
    $('#tableLink').attr('href', "");
    selection = null;
    currentTableId = null;
    tableLink = null;
    errorOccurred = false;
    // $('#display-data').text('You haven\'t selected any data yet.');
}
