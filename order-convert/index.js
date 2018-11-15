const XLSX = require('xlsx');

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');
        if (req.body.contents) {
            var workbook = XLSX.read((req.body.contents), {type : "base64", cellDates : true});
            console.log('error:', error); // Print the error if one occurred
            console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received
            var workbook = XLSX.read(body, {type : "buffer", cellDates : true});
            var wname = workbook.SheetNames[0];
			var worksheet = workbook.Sheets[wname];
          
            var formatMoney = function(amount, decimalCount = 2, decimal = ".", thousands = ",") {
                try {
                decimalCount = Math.abs(decimalCount);
                decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

                const negativeSign = amount < 0 ? "-" : "";

                let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
                let j = (i.length > 3) ? i.length % 3 : 0;

                return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
                } 
                catch (e) {
                    console.log(e)
                }
            };
          
            // Start from Excel row 2.
            var targetRows = [];
          	var row = 2;
            // Test column A until empty.
            var cell = worksheet["A" + row];
            var value = cell ? cell.v : "";
            var testValue = value;
            while (testValue != "") {
                var rowValues = {};
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
                rowValues["D"] = ("€ "+formatMoney(cell.v));
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

            // Create Excel in EasyFatt order format.
            var wsEasyFatt = XLSX.utils.json_to_sheet(
                targetRows/*,
                {
                    header : ["Cod.", "Descrizione", "Q.tà", "Prezzo netto", 
                        "U.m.", "Sconti", "IVA"
                    ],
                    skipHeader : true
                }*/
            )

            context.res = {
                // status: 200, /* Defaults to 200 */
                status : 201,
                body: {
                    wsEasyFatt
                },
                headers : {
                    "Content-Type" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Content-Disposition" : 'attachment;filename="Ordine EasyFatt.xlsx"'
                }
            };
        }
    else {
        context.res = {
            status: 400,
            body: "Please pass a filename and base64 contents in the request body"
        };
    }
};