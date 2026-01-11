function readSpreadsheet() {
  try {
    var sheet = SpreadsheetApp.openById("");
    var form = FormApp.openById("");

    if (sheet == null || form == null) {
      Logger.log("Spreadsheet or Form not found.");
      return;
    }

    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var values = range.getValues();
    var items = form.getItems();
    var MAX_ITEMS = 1000;
  
    // Check if first column is "timestamp" and should be skipped
    var headerRow = values[0];
    var skipFirstColumn = (headerRow[0] && headerRow[0].toString().toLowerCase().indexOf("timestamp") !== -1);
    
    // Maximum rows to process
    // Limit the number of rows to process to MAX_ITEMS or 500, whichever is smaller
    var maxRowsToProcess = Math.min(MAX_ITEMS, 500, numRows - 1);

    // row range to process
    var startRow = 1; // skip header row (0)
    var endRow = startRow + maxRowsToProcess - 1;

    Logger.log("Begin Processing Rows");

    for (var i = startRow; i <= endRow; i++) {
      try {
        var value = values[i];
        var formResponse = form.createResponse();
        var k = skipFirstColumn ? 1 : 0; // Skip timestamp column if header indicates it
        for (var j = 0; j < items.length; j++) {
          try {
            var item = items[j];
            var itemType = item.getType();
            switch (itemType) {
              case FormApp.ItemType.LIST:
                formResponse.withItemResponse(
                  item.asListItem().createResponse(value[k++])
                );
                break;
              case FormApp.ItemType.MULTIPLE_CHOICE:
                formResponse.withItemResponse(
                  item.asMultipleChoiceItem().createResponse(value[k++])
                );
                break;
              case FormApp.ItemType.PARAGRAPH_TEXT:
                formResponse.withItemResponse(
                  item.asParagraphTextItem().createResponse(value[k++])
                );
                break;
              case FormApp.ItemType.TEXT:
                formResponse.withItemResponse(
                  item.asTextItem().createResponse(value[k++])
                );
                break;
              case FormApp.ItemType.CHECKBOX:
                formResponse.withItemResponse(
                  item.asCheckboxItem().createResponse(
                    value[k++].split(",").map(function (option) {
                      return option.trim();
                    })
                  )
                );
                break;
              case FormApp.ItemType.SCALE:  
                formResponse.withItemResponse(
                  item.asScaleItem().createResponse(parseInt(value[k++]))
                );
                break;
              case FormApp.ItemType.GRID:
                var gridItem = item.asGridItem();
                var gridRows = gridItem.getRows();
                var gridResponses = [];
                for (var r = 0; r < gridRows.length; r++) {
                  if (value[k] !== "" && value[k] !== null && value[k] !== undefined) {
                    gridResponses.push(value[k]);
                  } else {
                    gridResponses.push(""); // or handle empty values as needed
                  }
                  k++;
                }
                if (gridResponses.length > 0) {
                  formResponse.withItemResponse(
                    gridItem.createResponse(gridResponses)
                  );
                }
                break;
              case FormApp.ItemType.PAGE_BREAK:
                break;
              default:
                Logger.log("Skipping unsupported item type: " + itemType);
            }
          } catch (itemError) {
            Logger.log("ItemError: Error processing item " + j + " in row " + i + ": " + itemError.message);
            Logger.log("itemType: " + (typeof itemType !== 'undefined' ? itemType : 'unknown'));
            Logger.log("Row data: " + JSON.stringify(values[i]));
            continue;// Continue to next item
          }
        }
        formResponse.submit();
      } catch (rowError) {
        Logger.log("RowError: Error processing row " + i + ": " + rowError.message);
        Logger.log("Row data: " + JSON.stringify(values[i]));
        continue;// Continue to next row
      }
    }
  } catch (error) {
    Logger.log("Error in readSpreadsheet: " + error.message);
  }
}
