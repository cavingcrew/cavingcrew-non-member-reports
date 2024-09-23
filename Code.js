const scriptProperties = PropertiesService.getScriptProperties();

const server = scriptProperties.getProperty('cred_server');
const port = parseInt(scriptProperties.getProperty('cred_port'), 10);
const dbName = scriptProperties.getProperty('cred_dbName');
const username = scriptProperties.getProperty('cred_username');
const password = scriptProperties.getProperty('cred_password');
const url = `jdbc:mysql://${server}:${port}/${dbName}`;
const cc_location = scriptProperties.getProperty('cred_cc_location');
const apidomain = scriptProperties.getProperty('cred_apidomain');
const apiusername = scriptProperties.getProperty('cred_apiusername');
const apipassword = scriptProperties.getProperty('cred_apipassword');
const reportYears = [2024,2023,2022,2021]

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Refresh Reports')
    .addItem('Run Main', 'main')
    .addToUi();
}

function main(){
let years = reportYears
writeToSpreadsheet(years)
}

function writeToSpreadsheet(years) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var connection = Jdbc.getConnection(url, username, password);

  years.forEach(function(year) {
    var sheet = spreadsheet.getSheetByName(year.toString());
    if (!sheet) {
      sheet = spreadsheet.insertSheet(year.toString());
    } else {
      // Clear the sheet contents and formatting before writing new data
      sheet.clear();
    }

var query = "SELECT m.first_name AS 'First Name', m.last_name AS 'Last Name', m.billing_address_1 AS 'Billing Address 1', m.billing_address_2 AS 'Billing Address 2', m.billing_city AS 'City', m.billing_postcode AS 'Postcode', trip1.trip_name AS 'Trip 1 Name', trip1.trip_date AS 'Trip 1 Date', trip2.trip_name AS 'Trip 2 Name', trip2.trip_date AS 'Trip 2 Date', trip3.trip_name AS 'Trip 3 Name', trip3.trip_date AS 'Trip 3 Date', trip4.trip_name AS 'Trip 4 Name', trip4.trip_date AS 'Trip 4 Date', m.id AS 'Caving Crew ID' FROM jtl_member_db m LEFT JOIN (SELECT * FROM (SELECT user_id, SUBSTRING_INDEX(order_item_name, ' - ', 1) AS trip_name, COALESCE(DATE(cc_start_date), DATE(order_created)) AS trip_date, ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY order_created) AS rn FROM jtl_order_product_customer_lookup) sub WHERE rn = 1) AS trip1 ON m.ID = trip1.user_id LEFT JOIN (SELECT * FROM (SELECT user_id, SUBSTRING_INDEX(order_item_name, ' - ', 1) AS trip_name, COALESCE(DATE(cc_start_date), DATE(order_created)) AS trip_date, ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY order_created) AS rn FROM jtl_order_product_customer_lookup) sub WHERE rn = 2) AS trip2 ON m.ID = trip2.user_id LEFT JOIN (SELECT * FROM (SELECT user_id, SUBSTRING_INDEX(order_item_name, ' - ', 1) AS trip_name, COALESCE(DATE(cc_start_date), DATE(order_created)) AS trip_date, ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY order_created) AS rn FROM jtl_order_product_customer_lookup) sub WHERE rn = 3) AS trip3 ON m.ID = trip3.user_id LEFT JOIN (SELECT * FROM (SELECT user_id, SUBSTRING_INDEX(order_item_name, ' - ', 1) AS trip_name, COALESCE(DATE(cc_start_date), DATE(order_created)) AS trip_date, ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY order_created) AS rn FROM jtl_order_product_customer_lookup) sub WHERE rn = 4) AS trip4 ON m.ID = trip4.user_id WHERE (m.cc_member IS NULL OR m.cc_member = '') AND EXISTS (SELECT 1 FROM jtl_order_product_customer_lookup WHERE user_id = m.ID AND YEAR(COALESCE(DATE(cc_start_date), DATE(order_created))) = " + year + " LIMIT 1) ORDER BY m.first_name, m.last_name;";

console.log("Getting results for ", year)
var results = connection.createStatement().executeQuery(query);


    var data = [];
    var headers = [];
    var row = 0;
    while (results.next()) {
      var record = [];
      if (row === 0) {

}
      for (var col = 0; col < results.getMetaData().getColumnCount(); col++) {
        if (row === 0) {
          headers.push(results.getMetaData().getColumnName(col + 1));
        }
        record.push(results.getString(col + 1));
      }
      data.push(record);
      row++;
    }

    results.close();

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  });

  connection.close();
}
