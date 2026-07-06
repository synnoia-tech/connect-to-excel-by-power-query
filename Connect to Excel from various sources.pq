let
  prExcelPath
    = "InputYourExcelPathHere",
  Source =
    if Text.Contains(
      prExcelPath,
      "https://docs.google.com/spreadsheets/d/"
    ) //Google Sheets is different than other cases                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
    then
      try //try with the public Google Sheet first
        Excel.Workbook(
          Web.Contents(
            "https://docs.google.com/spreadsheets/d/"
              & Text.BetweenDelimiters(prExcelPath, "/", "/", 4, 0)
              & "/export?format=xlsx&id="
              & Text.BetweenDelimiters(prExcelPath, "/", "/", 4, 0)
          ),
          null,
          true
        )
      otherwise //if there is error in the previous step, use the Beta GoogleSheets connector
        GoogleSheets.Contents(prExcelPath)
    else
      Excel.Workbook(Web.Contents(Replacer.ReplaceText(prExcelPath, "?web=1", "")), null, true) //use Web.Contents function for other cases, even the local files
in
  Source
