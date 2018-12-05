const XLSX = require('xlsx');

module.exports = async function (context, req) {
    try {
        debugger;
        context.log('JavaScript HTTP trigger function processed a request.');
        if (req.body.contents) {
            // Load file from body. Base64 encoded.
            try {
                var workbook = XLSX.read((req.body.contents), { type: "base64", cellDates: true });
                var wname = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[wname];
            }
            catch (e) {
                context.log(e);
                context.res = {
                    // status: 400, /* Error */
                    status: 400,
                    body: {
                        status: "error",
                        message: "Unable to load file.",
                        diagnostic: e.name + ' - ' + e.message
                    }
                };
                return;
            }

            // Check mode parameter to convert file or describe it.
            if (req.body.mode && req.body.mode == 'describe') {
                var cell = worksheet["AQ1"];
                var afn = cell ? cell.v : "";

                cell = worksheet["AR1"];
                var aln = cell ? cell.v : "";

                cell = worksheet["L1"];
                var cn = cell ? cell.v : "";

                context.res = {
                    status: 200,
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: {
                        "agent-first-name": afn,
                        "agent-last-name": aln,
                        "customer-name": cn,
                    }
                };

            } else {
                var formatMoney = function (amount, decimalCount = 2, decimal = ",", thousands = "") {
                    try {
                        decimalCount = Math.abs(decimalCount);
                        decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

                        const negativeSign = amount < 0 ? "-" : "";

                        let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
                        let j = (i.length > 3) ? i.length % 3 : 0;

                        return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
                    }
                    catch (e) {
                        context.log(e);
                        return;
                    }
                };

                try {
                    // Prepare header.
                    var targetRows = [];
                    var rowValues = {
                        A: "Cod.",
                        B: "Descrizione",
                        C: "Q.tà",
                        D: "Prezzo netto",
                        E: "U.m.",
                        F: "Sconti",
                        G: "Iva"
                    };
                    targetRows.push(rowValues);
                    // Start from Excel row 1.
                    var row = 1;
                    // Test column A until empty.
                    var cell = worksheet["A" + row];
                    var value = cell ? cell.v : "";
                    var testValue = value;
                    while (testValue != "") {
                        rowValues = {};
                        // Product code.
                        var cell = worksheet["AY" + row];
                        var value = cell ? cell.v : "";
                        rowValues["A"] = value;
                        // Product description.
                        var cell = worksheet["AZ" + row];
                        var value = cell ? cell.v : "";
                        rowValues["B"] = value;
                        // Quantity.	
                        var cell = worksheet["BA" + row];
                        var value = cell ? cell.v : "";
                        rowValues["C"] = value;
                        // Unit price.
                        var cell = worksheet["BB" + row];
                        var value = cell ? cell.v : "0";
                        //var value = cell ? ("€ "+Number(cell.v.toFixed(1)).toLocaleString()) : "";
                        rowValues["D"] = ("€ " + formatMoney(cell.v));
                        // Units name.
                        var cell = worksheet["BF" + row];
                        var value = cell ? cell.v : "";
                        rowValues["E"] = value;
                        // Discount percentage (basis points e.g. 20 = 20%).
                        var cell = worksheet["BP" + row];
                        var value = cell ? cell.v : undefined;
                        //If BP == 1 then F = 100
                        if (value == 1) {
                            rowValues["F"] = 100;
                            value = 100;
                        }
                        else {
                            var cell = worksheet["BC" + row];
                            var value = cell ? cell.v : "";
                            rowValues["F"] = value;
                        }
                        // VAT amount.
                        var cell = worksheet["BG" + row];
                        var value = cell ? cell.v : "";
                        rowValues["G"] = value;
                        // Append row at the bottom of the array.
                        targetRows.push(rowValues);
                        // Test column A until empty.
                        row++;
                        var cell = worksheet["A" + row];
                        var value = cell ? cell.v : "";
                        var testValue = value;
                    }
                }
                catch (e) {
                    context.log(e);
                    context.res = {
                        // status: 400, /* Error */
                        status: 400,
                        body: {
                            status: "error",
                            message: "Unable to convert file data.",
                            diagnostic: e.name + ' - ' + e.message
                        }
                    };
                    return;
                }

                // Create Excel workbook and append sheet 
                // in EasyFatt order format.
                var wsEasyFatt;
                var wbOutput = XLSX.utils.book_new();

                try {
                    wsEasyFatt = XLSX.utils.json_to_sheet(
                        targetRows,
                        { skipHeader: true }
                    )
                    XLSX.utils.book_append_sheet(wbOutput, wsEasyFatt, "Righe documento");
                    /* generate buffer */
                    var buf = XLSX.write(wbOutput, { type: 'buffer', bookType: "xlsx" });
                }
                catch (e) {
                    context.log(e);
                    context.res = {
                        // status: 400, /* Error */
                        status: 400,
                        body: {
                            status: "error",
                            message: "Unable to convert file data.",
                            diagnostic: e.name + ' - ' + e.message
                        }
                    };
                    return;
                }

                // Format response.
                context.res = {
                    status: 200,
                    headers: {
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Content-Disposition": 'attachment;filename="Ordine EasyFatt.xlsx"'
                    },
                    body: new Uint8Array(buf)/*,
            isRaw : true*/
                };
            }
        }
        else {
            context.res = {
                status: 400,
                body: "Please pass a filename and base64 contents in the request body"
            };
        }
    }
    catch (e) {
        //Error handling at global level
        context.res = {
            status: 400,
            body: e
        };
    }
};