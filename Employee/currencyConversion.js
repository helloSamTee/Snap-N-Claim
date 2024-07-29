function createCurrencyDropdown() {
  // List of currency ISO codes
  const currencyCodes = [
    "AED", "AFN", "ALL", "AMD", "ANG", "AOA", "ARS", "AUD", "AWG", "AZN",
    "BAM", "BBD", "BDT", "BGN", "BHD", "BIF", "BMD", "BND", "BOB", "BRL",
    "BSD", "BTN", "BWP", "BYN", "BZD", "CAD", "CDF", "CHF", "CLP", "CNY",
    "COP", "CRC", "CUC", "CUP", "CVE", "CZK", "DJF", "DKK", "DOP", "DZD",
    "EGP", "ERN", "ETB", "EUR", "FJD", "FKP", "FOK", "GBP", "GEL", "GGP",
    "GHS", "GIP", "GMD", "GNF", "GTQ", "GYD", "HKD", "HNL", "HRK", "HTG",
    "HUF", "IDR", "ILS", "IMP", "INR", "IQD", "IRR", "ISK", "JMD", "JOD",
    "JPY", "KES", "KGS", "KHR", "KID", "KMF", "KRW", "KWD", "KYD", "KZT",
    "LAK", "LBP", "LKR", "LRD", "LSL", "LYD", "MAD", "MDL", "MGA", "MKD",
    "MMK", "MNT", "MOP", "MRU", "MUR", "MVR", "MWK", "MXN", "MYR", "MZN",
    "NAD", "NGN", "NIO", "NOK", "NPR", "NZD", "OMR", "PAB", "PEN", "PGK",
    "PHP", "PKR", "PLN", "PYG", "QAR", "RON", "RSD", "RUB", "RWF", "SAR",
    "SBD", "SCR", "SDG", "SEK", "SGD", "SHP", "SLL", "SOS", "SRD", "SSP",
    "STN", "SYP", "SZL", "THB", "TJS", "TMT", "TND", "TOP", "TRY", "TTD",
    "TVD", "TWD", "TZS", "UAH", "UGX", "USD", "UYU", "UZS", "VES", "VND",
    "VUV", "WST", "XAF", "XCD", "XOF", "XPF", "YER", "ZAR", "ZMW", "ZWL"
  ];

  // Open the active spreadsheet and get the first sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notApproved");

  

  // Create a dropdown list for currency column (6)
  const range = sheet.getRange("H6");
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(currencyCodes).build();
  range.setDataValidation(rule);

}


function calculateMYR(cellRow, cellCol, sheet){
  const source = sheet.getRange(cellRow, cellCol).getValue();
  const total = sheet.getRange(cellRow, 9).getA1Notation();
  var formula;
  
  if (source == "MYR"){
    formula = `= ${total}`;
  }else{
    const date = sheet.getRange(cellRow, 4).getA1Notation();
    formula = `= ${total} * (index(GOOGLEFINANCE("CURRENCY:${source}MYR", "price", ${date}), 2, 2))`;
  }
  
  // Set formula to the cell
  sheet.getRange(cellRow, 10).setFormula(formula);
  // Example formula
  // =GOOGLEFINANCE(“CURRENCY:USDGBP”, “price”, DATE(2023,1,1), DATE(2023,1,1), “DAILY”)
}
