/*jshint esversion: 8 */

const request = require('request');
const xml2js = require('xml2js');
const dateformat = require("dateformat");
const parser = new xml2js.Parser({attrkey: 'Attr'});
const createCsvStringifier = require('csv-writer').createObjectCsvStringifier;

module.exports = (context, req) => {
    context.log('JavaScript HTTP trigger function processed a request.');

    //const {sourceUrl, expand, outFilename} = req.query || req.body;   d
    sourceUrl = req.body.sourceUrl;
    expand = req.body.expand;
    outFilename = req.body.outFilename;
    reportDate = req.body.reportDate;
    reportTime = req.body.reportTime;

    
    //set default report time to current date
    if(reportDate ==undefined)  {
              
        reportDate = dateformat(new Date(),"UTC:yyyy-mm-dd");
    }

    //set default report time  to current utc time
    if(reportTime ==undefined) {
      
        reportTime = dateformat(new Date(),"UTC:yyyy-mm-dd'T'hh:MM:ss'Z'");
    }   
    //return if expand is not rows or payments
    if(!(expand === 'rows' || expand === 'payments')) {
        context.res = {
            status: 400,
            body: "Parameter 'expand' should be passed with either 'rows' or 'payments'"
        };
        context.done();
        return;
    }
   

    getOptions = {
        method: 'GET',
        uri: sourceUrl,
    };

    request(getOptions, (err, res, req) => {
        if (err) {
            context.res = {
                status: res.statusCode,
                body: err,
            };
        } else {
            parser.parseString(req, (err, res) => {
                if (!err) {
                    let temp = res.EasyfattDocuments.Company[0];

                    //set Company data
                    let company = {};
                   
                    
                    setColumnHeaders(company);

                    for (const key in temp) {
                        if (temp.hasOwnProperty(key)) {
                            company['Company.' + key] = temp[key][0];
                        }
                    }
                    //set report Date and report Time
                    company["ReportDate"] = reportDate;
                    company["ReportTime"] = reportTime;
                    temp = res.EasyfattDocuments.Documents[0].Document;
                    let documents = [];
                    let orderID = '', orderDate = '', shipID = '', shipDate = '';

                    try{

                        for (let index = 0; index < temp.length; index++) {
                            const item = temp[index];
                            let newEl ={};
                            
                            
                            Object.assign(newEl,company);
                            
                            let childItems = [];
                            //This is the processing of documents. Iterate each document 
                            //var item is representing the documents element
                            for (const key in item) {
                                if (item.hasOwnProperty(key)) {
                                    switch (key) {
                                        case 'CostVatCode':
                                            newEl['Document.' + key] = item[key][0]._;
                                            break;
                                        case 'Payments':
                                            //This case will executed when expand is set to "payments" as paramenter
                                            if (expand === 'payments') {
                                                if (item[key][0].hasOwnProperty('Payment')) {
                                                    childItems = item[key][0].Payment;
                                                    for (let childIndex = 0; childIndex < childItems.length; childIndex++) {
                                                        let row = Object.assign({}, newEl);
                                                        const childItem = childItems[childIndex];
                                                        //set Paymenet details
                                                        row['PaymentRow.RowNumber'] = childIndex + 1;

                                                        var paddedRowId = pad(childIndex +1,4);
                                                        row['PaymentRow.RowID'] = "P" + "|"
                                                                                    + item["DocumentType"] + "|"
                                                                                    + item["Date"] + "|"
                                                                                    + item["Number"] + "|"
                                                                                    + item["Numbering"] + "|"
                                                                                    + paddedRowId;
                                                        for (const childKey in childItem) {
                                                            if (childItem.hasOwnProperty(childKey)) {
                                                                row['PaymentRow.' + childKey] = childItem[childKey][0];
                                                            }
                                                        }
                                                        documents.push(row);
                                                    }
                                                } else {
                                                    documents.push(newEl);
                                                }
                                            }
                                            break;
                                        case 'Rows':
                                            //This case will executed when expand is set to "rows" as paramenter
                                            if (expand === 'rows') {
                                                if (item[key][0].hasOwnProperty('Row')) {                                                   
                                                    childItems = item[key][0].Row;
                                                     //Reset OrderId, OrderDate, ShipId,shipDate on new rows element starts.
                                                     //This is the process of rows  element.
                                                     let orderID = '', orderDate = '', shipID = '', shipDate = '';
                                                    for (let childIndex = 0; childIndex < childItems.length; childIndex++) {
                                                        let row = Object.assign({}, newEl);
                                                        const childItem = childItems[childIndex];
                                                        //set Document Details
                                                        row['DocumentRow.RowNumber'] = childIndex + 1;
                                                        var paddedRowId = pad(childIndex +1,4);
                                                        row['DocumentRow.RowID'] = "R" + "|"
                                                                                    + item["DocumentType"] + "|"
                                                                                    + item["Date"] + "|"
                                                                                    + item["Number"] + "|"
                                                                                    + item["Numbering"] + "|"
                                                                                    + paddedRowId;
                                                        //This is the process of rows under each 
                                                        for (const childKey in childItem) {
                                                            if (childItem.hasOwnProperty(childKey)) {
                                                                if (childKey === 'VatCode') {
                                                                    if (childItem[childKey][0]._ === undefined) row['DocumentRow.' + childKey] = "";
                                                                    else row['DocumentRow.' + childKey] = childItem[childKey][0]._;
                                                                } else {
                                                                    row['DocumentRow.' + childKey] = childItem[childKey][0];
                                                                }
                                                            }
                                                        }
                                                        if (childItem.Description[0].startsWith("** Rif. Conferma d'ordine")) {
                                                            let uDesc = childItem.Description[0].replace("** Rif. Conferma d'ordine","");
                                                            uDesc = uDesc.trim();
                                                            let [id, date] = uDesc.split(" del ");
                                                            date = date.replace(":", "");
                                                            let dates = date.split("/");
                                                            date = "";
                                                            for(let ind = dates.length - 1; ind >= 0; ind --)
                                                                date += dates[ind] + '-';
                                                            date = date.slice(0, -1);
                                                            orderID = id;
                                                            orderDate = date;
                                                        } else if (childItem.Description[0].startsWith("** Rif. Doc. di trasporto")) {
                                                            let uDesc = childItem.Description[0].replace("** Rif. Doc. di trasporto","");
                                                            uDesc = uDesc.trim();
                                                            let [id, date] = uDesc.split(" del ");
                                                            date = date.replace(":", "");
                                                            let dates = date.split("/");
                                                            date = "";
                                                            for(let ind = dates.length - 1; ind >= 0; ind --)
                                                                date += dates[ind] + '-';
                                                            date = date.slice(0, -1);
                                                            shipID = id;
                                                            shipDate = date;
                                                            //Reset Orderid and OrderDate when new shipment found
                                                            orderID = "";
                                                            orderDate = "";
                                                        }
                                                        row['DocumentRow.OrderID'] = orderID;
                                                        row['DocumentRow.OrderDate'] = orderDate;
                                                        row['DocumentRow.ShipmentID'] = shipID;
                                                        row['DocumentRow.ShipmentDate'] = shipDate;
                                                        documents.push(row);
                                                    }
                                                } else {
                                                    documents.push(newEl);
                                                }
                                            }
                                            break;
                                        default:
                                            //This is for invoice data
                                            newEl['Document.' + key] = item[key][0];
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    catch(e)
                    {
                        console.log("Error: " + e);
                    }
                   
                    const csvHeader = Object.keys(company).map((value) => {
                            return {
                                'id': value,
                                'title': value,
                            };
                        });
                        
                    
                    const csvStringifier = createCsvStringifier({
                        header: csvHeader,
                    });
                    
                    

                    let bufStr = csvStringifier.getHeaderString();
                    bufStr += csvStringifier.stringifyRecords(documents);
                    const buf = Buffer.from(bufStr);

                    context.res = {
                        status: 200,
                        headers: {
                            "Content-Type": "text/csv",
                            "Content-Disposition": 'attachment;filename="'+outFilename+'"'
                        },
                        body: new Uint8Array(buf)
                    };
                    context.done();
                } else {
                    context.res = {
                        status: res.statusCode,
                        body: err,
                    };
                    context.done();
                }
            });
        }
    });


    // fs.readFile("24.05.19 ddt completi.DefXml", (err, data) => {
    //     if (err) {
    //         context.res = {
    //             status: 500,
    //             body: err,
    //         };
    //     } else {
    //         parser.parseString(data, (err, res) => {
    //             if (!err) {
    //                 let temp = res.EasyfattDocuments.Company[0];

    //                 let company = {};
    //                 for (const key in temp) {
    //                     if (temp.hasOwnProperty(key)) {
    //                         company['Company.' + key] = temp[key][0];
    //                     }
    //                 }

    //                 temp = res.EasyfattDocuments.Documents[0].Document;
    //                 let documents = [];
    //                 let orderID = '', orderDate = '', shipID = '', shipDate = '';

    //                 for (let index = 0; index < temp.length; index++) {
    //                     const item = temp[index];
    //                     let newEl = Object.assign({}, company);
    //                     let childItems = [];
    //                     for (const key in item) {
    //                         if (item.hasOwnProperty(key)) {
    //                             switch (key) {
    //                                 case 'CostVatCode':
    //                                     newEl['Document.' + key] = item[key][0]._;
    //                                     break;
    //                                 case 'Payments':
    //                                     if (expand === 'payments') {
    //                                         if (item[key][0].hasOwnProperty('Payment')) {
    //                                             childItems = item[key][0].Payment;
    //                                             for (let childIndex = 0; childIndex < childItems.length; childIndex++) {
    //                                                 let row = Object.assign({}, newEl);
    //                                                 const childItem = childItems[childIndex];
    //                                                 row['PaymentRow.RowNumber'] = childIndex + 1;
    //                                                 row['PaymentRow.RowID'] = "P" + "|"
    //                                                                             + item["DocumentType"] + "|"
    //                                                                             + item["Date"] + "|"
    //                                                                             + item["Number"] + "|"
    //                                                                             + item["Numbering"] + "|"
    //                                                                             + childIndex;
    //                                                 for (const childKey in childItem) {
    //                                                     if (childItem.hasOwnProperty(childKey)) {
    //                                                         row['PaymentRow.' + childKey] = childItem[childKey][0];
    //                                                     }
    //                                                 }
    //                                                 documents.push(row);
    //                                             }
    //                                         } else {
    //                                             documents.push(newEl);
    //                                         }
    //                                     }
    //                                     break;
    //                                 case 'Rows':
    //                                     if (expand === 'rows') {
    //                                         if (item[key][0].hasOwnProperty('Row')) {
    //                                             childItems = item[key][0].Row;
    //                                             for (let childIndex = 0; childIndex < childItems.length; childIndex++) {
    //                                                 let row = Object.assign({}, newEl);
    //                                                 const childItem = childItems[childIndex];
    //                                                 row['DocumentRow.RowNumber'] = childIndex + 1;
    //                                                 row['DocumentRow.RowID'] = "R" + "|"
    //                                                                             + item["DocumentType"] + "|"
    //                                                                             + item["Date"] + "|"
    //                                                                             + item["Number"] + "|"
    //                                                                             + item["Numbering"] + "|"
    //                                                                             + childIndex;
    //                                                 for (const childKey in childItem) {
    //                                                     if (childItem.hasOwnProperty(childKey)) {
    //                                                         if (childKey === 'VatCode') {
    //                                                             if (childItem[childKey][0]._ === undefined) row['DocumentRow.' + childKey] = "";
    //                                                             else row['DocumentRow.' + childKey] = childItem[childKey][0]._;
    //                                                         } else {
    //                                                             row['DocumentRow.' + childKey] = childItem[childKey][0];
    //                                                         }
    //                                                     }
    //                                                 }
    //                                                 if (childItem.Description[0].startsWith("** Rif. Conferma d'ordine")) {
    //                                                     let uDesc = childItem.Description[0].replace("** Rif. Conferma d'ordine","");
    //                                                     uDesc = uDesc.trim();
    //                                                     let [id, date] = uDesc.split(" del ");
    //                                                     date = date.replace(":", "");
    //                                                     let dates = date.split("/");
    //                                                     date = "";
    //                                                     for(let ind = dates.length - 1; ind >= 0; ind --)
    //                                                         date += dates[ind] + '-';
    //                                                     date = date.slice(0, -1);
    //                                                     orderID = id;
    //                                                     orderDate = date;
    //                                                 } else if (childItem.Description[0].startsWith("** Rif. Doc. di trasporto")) {
    //                                                     let uDesc = childItem.Description[0].replace("** Rif. Doc. di trasporto","");
    //                                                     uDesc = uDesc.trim();
    //                                                     let [id, date] = uDesc.split(" del ");
    //                                                     date = date.replace(":", "");
    //                                                     let dates = date.split("/");
    //                                                     date = "";
    //                                                     for(let ind = dates.length - 1; ind >= 0; ind --)
    //                                                         date += dates[ind] + '-';
    //                                                     date = date.slice(0, -1);
    //                                                     shipID = id;
    //                                                     shipDate = date;
    //                                                 }
    //                                                 row['DocumentRow.OrderID'] = orderID;
    //                                                 row['DocumentRow.OrderDate'] = orderDate;
    //                                                 row['DocumentRow.ShipmentID'] = shipID;
    //                                                 row['DocumentRow.ShipmentDate'] = shipDate;
    //                                                 documents.push(row);
    //                                             }
    //                                         } else {
    //                                             documents.push(newEl);
    //                                         }
    //                                     }
    //                                     break;
    //                                 default:
    //                                     newEl['Document.' + key] = item[key][0];
    //                                     break;
    //                             }
    //                         }
    //                     }
    //                 }

    //                 const csvHeader = Object.keys(documents[0]).map((value) => {
    //                     return {
    //                         'id': value,
    //                         'title': value,
    //                     };
    //                 });

    //                 const csvWriter = createCsvWriter({
    //                     path: 'output.csv',
    //                     header: csvHeader,
    //                 });

    //                 csvWriter.writeRecords(documents)
    //                 .then(() => {
    //                     fs.readFile('output.csv', function(err,data){
    //                         if (!err) {
    //                             context.res = {
    //                                 status: 200,
    //                                 headers: {
    //                                     "Content-Type": "text/csv",
    //                                     "Content-Disposition": 'attachment;filename="'+outFilename+'"'
    //                                 },
    //                                 body: new Uint8Array(data)
    //                             };
    //                             context.done();
    //                         } else {
    //                             context.res = {
    //                                 status: 500,
    //                                 body: err.toString(),
    //                             };
    //                             context.done();
    //                         }
    //                     });
    //                 })
    //                 .catch((res) => {
    //                     context.res = {
    //                         status: 500,
    //                         body: res.toString(),
    //                     };
    //                     context.done();
    //                 });
    //             } else {
    //                 context.res = {
    //                     status: res.statusCode,
    //                     body: err,
    //                 };
    //                 context.done();
    //             }
    //         });
    //     }
    // });
};

//This function function will return padded character passed as 'z'. Default value of 'z' is '0'
function pad(n, width, z) {
    z = z || '0';
    n = n + '';
    return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
  }

function setColumnHeaders(headers) {
    headers["ReportDate"] = '';
    headers["ReportTime"] = '';
    headers['Company.Name'] = '';
    headers['Company.Address'] = '';
    headers['Company.Postcode'] = '';
    headers['Company.City'] = '';
    headers['Company.Province'] = '';
    headers['Company.Country'] = '';
    headers['Company.FiscalCode'] = '';
    headers['Company.VatCode'] = '';
    headers['Company.Tel'] = '';
    headers['Company.Email'] = '';
    headers['Company.HomePage'] = '';
    headers['Document.Date'] = '';
    headers['Document.Number'] = '';
    headers['Document.Numbering'] = '';
    headers['Document.CustomerName'] = '';
    headers['Document.CustomerCode'] = '';
    headers['Document.CustomerWebLogin'] = '';
    headers['Document.CustomerAddress'] = '';
    headers['Document.CustomerPostcode'] = '';
    headers['Document.CustomerCity'] = '';
    headers['Document.CustomerProvince'] = '';
    headers['Document.CustomerCountry'] = '';
    headers['Document.CustomerFiscalCode'] = '';
    headers['Document.CustomerVatCode'] = '';
    headers['Document.CustomerReference'] = '';
    headers['Document.CustomerTel'] = '';
    headers['Document.CustomerCellPhone'] = '';
    headers['Document.CustomerFax'] = '';
    headers['Document.CustomerEmail'] = '';
    headers['Document.CustomerPec'] = '';
    headers['Document.CustomerEInvoiceDestCode'] = '';
    headers['Document.DocumentType'] = '';
    headers['Document.DeliveryName'] = '';
    headers['Document.DeliveryAddress'] = '';
    headers['Document.DeliveryPostcode'] = '';
    headers['Document.DeliveryCity'] = '';
    headers['Document.DeliveryProvince'] = '';
    headers['Document.DeliveryCountry'] = '';
    headers['Document.Warehouse'] = '';
    headers['Document.CostDescription'] = '';
    headers['Document.CostVatCode'] = '';
    headers['Document.CostAmount'] = '';
    headers['Document.TotalWithoutTax'] = '';
    headers['Document.VatAmount'] = '';
    headers['Document.TotalSubjectToWithholdingTax'] = '';
    headers['Document.WithholdingTaxAmount'] = '';
    headers['Document.WithholdingTaxAmountB'] = '';
    headers['Document.Total'] = '';
    headers['Document.PriceList'] = '';
    headers['Document.PricesIncludeVat'] = '';
    headers['Document.WithholdingTaxPerc'] = '';
    headers['Document.WithholdingTaxPerc2'] = '';
    headers['Document.WithholdingTaxNameB'] = '';
    headers['Document.ContribDescription'] = '';
    headers['Document.ContribPerc'] = '';
    headers['Document.ContribSubjectToWithholdingTax'] = '';
    headers['Document.ContribAmount'] = '';
    headers['Document.ContribVatCode'] = '';
    headers['Document.PaymentName'] = '';
    headers['Document.PaymentBank'] = '';
    headers['Document.CustomField1'] = '';
    headers['Document.CustomField2'] = '';
    headers['Document.CustomField3'] = '';
    headers['Document.CustomField4'] = '';
    headers['Document.FootNotes'] = '';
    headers['Document.SalesAgent'] = '';
    headers['Document.DelayedVat'] = '';
    headers['Document.DelayedVatDesc'] = '';
    headers['Document.DelayedVatDueWithinOneYear'] = '';
    headers['Document.PaymentAdvanceAmount'] = '';
    headers['Document.Carrier'] = '';
    headers['Document.TransportReason'] = '';
    headers['Document.GoodsAppearance'] = '';
    headers['Document.TransportDateTime'] = '';
    headers['Document.ShipmentTerms'] = '';
    headers['Document.TransportedWeight'] = '';
    headers['Document.TrackingNumber'] = '';
    headers['Document.InternalComment'] = '';
    headers['Document.NumOfPieces'] = '';
    headers['Document.ExpectedConclusionDate'] = '';
    headers['Document.DocReference'] = '';
    headers['Document.Pdf'] = '';
    headers['DocumentRow.RowNumber'] = '';
    headers['DocumentRow.RowID'] = '';
    headers['DocumentRow.Code'] = '';
    headers['DocumentRow.Description'] = '';
    headers['DocumentRow.Qty'] = '';
    headers['DocumentRow.Um'] = '';
    headers['DocumentRow.Price'] = '';
    headers['DocumentRow.Discounts'] = '';
    headers['DocumentRow.VatCode'] = '';
    headers['DocumentRow.Total'] = '';
    headers['DocumentRow.Stock'] = '';
    headers['DocumentRow.Notes'] = '';
    headers['DocumentRow.OrderID'] = '';
    headers['DocumentRow.ShipmentID'] = '';
    headers['DocumentRow.ShipmentDate'] = '';
    headers['DocumentRow.OrderDate'] = '';
   
    headers['PaymentRow.RowNumber'] = '';
    headers['PaymentRow.RowID'] = '';
    headers['PaymentRow.Advance'] = '';
    headers['PaymentRow.Date'] = '';
    headers['PaymentRow.Amount'] = '';
    headers['PaymentRow.Paid'] = '';
}
