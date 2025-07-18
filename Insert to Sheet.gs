/**
 * Universal API Response Handler.
 *
 * This file processes and maps responses from all API endpoints. It replaces the
 * previous, impractical method of using a dedicated file for each process,
 * which is not scalable for the ~80 endpoints in the new API.
 */

function handleAccountTotals(response, sheetName) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    }

    let headers = [
      [
        'Property',
        'Property Name',
        'Property Id',
        'Property Address',
        'Property Street',
        'Property Street2',
        'Property City',
        'Property State',
        'Property Zip',
        'Reserve Amount',
        'Net Amount',
        'Ending Balance',
      ]
    ]
    sheet.getRange(8, 1, 1, headers[0].length).setValues(headers);

    let output = [];

    response.forEach(function (element) {
      output.push(
        [
          element["property"],
          element["property_name"],
          element["property_id"],
          element["property_address"],
          element["property_street"],
          element["property_street2"],
          element["property_city"],
          element["property_state"],
          element["property_zip"],
          element["reserve_amount"],
          element["net_amount"],
          element["ending_balance"],
        ]);
    });

    if (output.length === 0) {
      output.push(["No data returned!"]);
    }

    sheet.getRange(9, 1, output.length, output[0].length).setValues(output);
    Logger.log('New records entered.');
}
