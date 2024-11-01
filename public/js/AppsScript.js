const baseURL = "{site_url}",
  accessToken = "{token}",
  sheetTab = "{sheet_tab}";


  function RunSSGSW() {
    if (!ssgsw_current_sheet()) return;
    ssgsw_add_menus();
    ssgsw_apply_styles();
}

function onOpen() {
    RunSSGSW();
}

function ssgsw_current_sheet() {
    return SpreadsheetApp.getActiveSheet().getSheetName() == sheetTab;
}

function ssgsw_check_other_tab_edit(e) {
    var sheet = e.source.getActiveSheet();
    return sheet.getName();
}

function onEdit(e) {
    if (e == null || e == 'undefined' || e == '') return;
    if (e.triggerUid == null) return;
    if (ssgsw_check_other_tab_edit(e) === sheetTab) {
        const currentColumn = ssgsw_current_column();
        var ssgs_ui = SpreadsheetApp.getUi();
        if (["ID", "type", "sales", "category", "attributes","ssgsw_product_url"].includes(currentColumn)) {
            ssgs_ui.alert("This column is not editable");
            ssgsw_fetch_from_WordPress('manual');
            return;
        }
        ssgsw_toast2('Updating product... Please wait!.', 'Loading...');
        let data = ssgsw_get_edited_data(e);
        let key_value = ssgsw_columns(currentColumn, true);
        if (typeof key_value === "undefined") {
            key_value = currentColumn;
        }
        if (typeof key_value === 'object' && key_value !== null) {
            ssgs_ui.alert("This column is not editable");
            ssgsw_fetch_from_WordPress('manual');
        }
        let message = key_value + " updated successfully!";
        if (data[0]['ID'] == "") {
            message = "Product created successfully!";
        }
        ssgsw_sync_data(data, message, e);
    } else {
        var ssgsw_editedRange_other = e.range; 
        var editedCell_nation_other = ssgsw_editedRange_other.getA1Notation();
        var range_other = editedCell_nation_other.split(":");
        var startCell_other = range_other[0];
        var endCell_other = range_other[1] || range_other[0];
        var startRow_other = parseInt(startCell_other.substring(1));
        var endRow_other = parseInt(endCell_other.substring(1));
        var startColumn_other = startCell_other.charAt(0);
        var references_data = [];
        if (range_other[1] === undefined) {
            references_data.push(ssgsw_reference_row_data(e, startColumn_other + startRow_other));
        } else {
            for (var i = startRow_other; i <= endRow_other; i++) {
                references_data.push(ssgsw_reference_row_data(e, startColumn_other + i));
            }
        }
        var all_edit_data = ssgsw_format3(references_data[0]);
        let message = "Formula data updated successfully!";
        if (all_edit_data.length === 0) {
            return;
        }
        ssgsw_toast2('Updating product... Please wait!.', 'Loading...');
        ssgsw_sync_data(all_edit_data, message, e);
    }
}

function ssgsw_reference_row_data(e, startColumn_otherTest) {
    var all_data = ssgsw_reference_row_data_same_tab(e, startColumn_otherTest);
    var parent_sheet = e.source.getSheetByName(sheetTab);
    var sheet1Range = parent_sheet.getDataRange();
    var formulas = sheet1Range.getFormulas();
    var rowsWithReferences = [];
    all_data.forEach(function(charNation) {
        for (var i = 1; i <= formulas.length; i++) {
          for (var j = 1; j <= formulas[0].length; j++) {
            var formula = formulas[i - 1][j - 1];
            if (formula && formula.indexOf(ssgsw_check_other_tab_edit(e)) != -1) {
              if (formula && formula.indexOf(':') !== -1) {
                var check_formula_exits = ssgsw_formula_processDynamicText(formula, charNation);
                if (check_formula_exits) {
                    var referencedRow = sheet1Range.getCell(i, j).getRow();
                    var rowValues_range = parent_sheet.getRange(referencedRow, 1, 1, parent_sheet.getLastColumn());
                    var get_values = rowValues_range.getValues();
                    var get_formula = rowValues_range.getFormulas();
                    var marge_image_url = merge_and_parse_image_urls(get_values, get_formula)[0];
                    marge_image_url.unshift(referencedRow);
                    rowsWithReferences.push(marge_image_url);
                }
              } else if (formula && formula.indexOf(charNation) != -1) {
                var referencedRow = sheet1Range.getCell(i, j).getRow();
                var rowValues_range = parent_sheet.getRange(referencedRow, 1, 1, parent_sheet.getLastColumn());
                var get_values = rowValues_range.getValues();
                var get_formula = rowValues_range.getFormulas();
                var marge_image_url = merge_and_parse_image_urls(get_values, get_formula)[0];
                marge_image_url.unshift(referencedRow);
                rowsWithReferences.push(marge_image_url);
              }
            }
          }
        }
    });
    return rowsWithReferences;
}

function ssgsw_reference_row_data_same_tab(e, startColumn_otherTest) {
    var path = ssgsw_find_formula_relations(e, startColumn_otherTest);
    return path;
}

function ssgsw_find_formula_relations(e, startCell) {
    var visited = {};
    var paths = [];
    var activeSheet = e.source.getActiveSheet();
    var sheet1Range = activeSheet.getDataRange();
    var formulas = sheet1Range.getFormulas();
    function ssgsw_explore(cell, path) {
        var cellAddress = cell;
        if (visited[cellAddress]) {
          return;
        }
        visited[cellAddress] = true;
        path.push(cellAddress);
        paths.push(cellAddress);
        for (var i = 1; i <= formulas.length; i++) {
          for (var j = 1; j <= formulas[0].length; j++) {
            var formula = formulas[i - 1][j - 1];
            if (formula && formula.indexOf(':') !== -1) {
                var check_formula_exits = ssgsw_formula_processDynamicText(formula, cellAddress);
                if (check_formula_exits) {
                  var referencedRow = sheet1Range.getCell(i, j);
                  var referencedCellA1Notation = referencedRow.getA1Notation();
                  ssgsw_explore(referencedCellA1Notation,path);
                }
            } else if(formula && formula.indexOf(cellAddress) !== -1) {
              var referencedRow = sheet1Range.getCell(i, j);
              var referencedCellA1Notation = referencedRow.getA1Notation();
              ssgsw_explore(referencedCellA1Notation,path);
            }
          }
        }
    }
    ssgsw_explore(startCell, []);
    return paths;
  }

function ssgsw_formula_processDynamicText(text,referenceCell) {
    var texts = [
      text
    ];
    var cells = [];
    texts.forEach(function(text) {
      cells = cells.concat(ssgsw_formula_extractCellReferences(text));
    });
    var columns = [];
    var rows = [];
    cells.forEach(function(cell) {
      var split = ssgsw_formula_splitCell(cell);
      columns.push(split[0]);
      if (split[1]) {
        rows.push(parseInt(split[1]));
      }
    });
    var smallestColumn = ssgsw_formula_findSmallestColumn(columns);
    var largestColumn = ssgsw_formula_findLargestColumn(columns);
    var smallestRow = rows.length > 0 ? Math.min.apply(null, rows) : null;
    var largestRow = rows.length > 0 ? Math.max.apply(null, rows) : null;
    var referenceSplit = ssgsw_formula_splitCell(referenceCell);
    var referenceColumn = referenceSplit[0];
    var referenceRow = parseInt(referenceSplit[1]);
    return ssgsw_formula_compare(referenceColumn, referenceRow, smallestColumn, smallestRow, largestColumn, largestRow);
  }
  
  function ssgsw_formula_extractCellReferences(text) {
    var regex = /(?:[A-Za-z0-9]+!)?([A-Z]+[0-9]*)(?=:|$|,|\))/g;
    var matches = [];
    var match;
    while ((match = regex.exec(text)) !== null) {
      matches.push(match[1]);
    }
    return matches;
  }
  
  function ssgsw_formula_splitCell(cell) {
    var match = cell.match(/([A-Z]+)([0-9]*)/);
    return [match[1], match[2]];
  }
  
  function ssgsw_formula_findSmallestColumn(columns) {
    return columns.reduce(function(a, b) {
      return a < b ? a : b;
    });
  }
  
  function ssgsw_formula_findLargestColumn(columns) {
    return columns.reduce(function(a, b) {
      return a > b ? a : b;
    });
  }
  function ssgsw_formula_compare(referenceColumn, referenceRow, smallestColumn, smallestRow, largestColumn, largestRow) {
    if ( smallestRow == largestRow ) {
      var isUnderColumn = referenceColumn >= smallestColumn && referenceColumn <= largestColumn;
      return isUnderColumn;
    } else {
      var isUnderColumn = referenceColumn >= smallestColumn && referenceColumn <= largestColumn;
      var isUnderRow = referenceRow >= smallestRow && referenceRow <= largestRow;
      return isUnderColumn && isUnderRow;
    }
  
  }

function ssgsw_add_menus() {
    SpreadsheetApp.getUi().createMenu("Stock Sync").addItem("⟱ Fetch from WooCommerce", "ssgsw_fetch_from_WordPress").addItem("⟰ Sync on WooCommerce", "ssgsw_sync_all").addSeparator().addItem(" Format Styles", "ssgsw_apply_styles").addItem(" About Stock Sync", "ssgsw_about_us").addToUi();
}

function ssgsw_apply_styles() {
    const Columns = ssgsw_column_char(ssgsw_max_column());
    const Headers = "A1:" + Columns + 1;
    const StaticColumns = ssgsw_column_index(["ID", "type", "sales", "attributes", "category",'ssgsw_product_url']).filter((index) => index >= 0).map(ssgsw_column_char);
    Logger.log(StaticColumns);
    const StockColumn = ssgsw_column_index(["stock"]).map((index) => ssgsw_column_char(index))[0];
    const StaticColumnHeaders = StaticColumns.map((column) => column + 1);
    const StaticColumnValues = StaticColumns.map((column) => column + 2 + ":" + column);
    const StockColumnValues = StockColumn + 2 + ":" + StockColumn;
    const CenterableColumns = ssgsw_column_index(["stock", "regular_price", "sale_price", "sales", ]).filter((index) => index >= 0).map((char) => ssgsw_column_char(char) + "1:" + ssgsw_column_char(char));
    const Color = {
        primary: "#686de0",
        white: "white",
        black: "black",
        grey: "#dedede",
        success: "green",
        error: "indianred",
        info: "purple",
        warning: "orange",
    };
    const CurrentSheet = SpreadsheetApp.getActive().getSheetByName(sheetTab);
    CurrentSheet.getRange("A1:Z1").setFontWeight("normal").setBackground(Color.white).setFontColor(Color.black);
    CurrentSheet.getRange(Headers).setFontWeight("bold").setBackground(Color.primary).setFontColor(Color.white);
    CurrentSheet.autoResizeColumns(1, ssgsw_max_column());
    CurrentSheet.getRangeList(StaticColumnHeaders).setBackground(Color.error);
    CurrentSheet.getRangeList(StaticColumnValues).setBackground(Color.grey).setFontColor(Color.black);
    CurrentSheet.getRangeList(CenterableColumns).setHorizontalAlignment("center");
    let rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("In Stock").setBackground("#f7fff9").setFontColor("green").setRanges([SpreadsheetApp.getActiveSheet().getRange(StockColumnValues)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Out of Stock").setBackground("#fff8f7").setFontColor(Color.error).setRanges([SpreadsheetApp.getActiveSheet().getRange(StockColumnValues)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("On Backorder").setBackground("#fffdf7").setFontColor("orange").setRanges([SpreadsheetApp.getActiveSheet().getRange(StockColumnValues)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#f7fff9").setFontColor(Color.success).setRanges([SpreadsheetApp.getActiveSheet().getRange(StockColumnValues)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setBackground("#fff8f7").setFontColor(Color.error).setRanges([SpreadsheetApp.getActiveSheet().getRange(StockColumnValues)]).build());
    if (ssgsw_column_index(["sales"]) >= 0) {
        const SaleColumn = ssgsw_column_index(["sales"]).map((index) => ssgsw_column_char(index))[0];
        const SaleColumnValues = SaleColumn + 2 + ":" + SaleColumn;
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setFontColor(Color.error).setRanges([SpreadsheetApp.getActiveSheet().getRange(SaleColumnValues)]).build());
    }
    CurrentSheet.setConditionalFormatRules(rules);
    Logger.log("Stock Sync Initialized!");
}

function ssgsw_about_us() {
    let htmlOutput = HtmlService.createHtmlOutput(`<h3>Stock Sync with Google Sheet for WooCommerce</h3> <p>Sync your WooCommerce product stock with Google Sheets.</p> <p><a href="https://wordpress.org/plugins/stock-sync-with-google-sheet-for-woocommerce/" target="_blank">Download Free</a> version from WordPress.org</p> <p><a href="http://wppool.dev/stock-sync-for-woocommerce-with-google-sheet" target="_blank">Get Ultimate</a> version to enjoy all premium features and official updates.</p> `).setWidth(550).setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Stock Sync with Google Sheet for WooCommerce");
}

function ssgsw_api_access_block_notice() {
    let htmlOutput = HtmlService.createHtmlOutput(`<p><strong>Oopss!</strong> It looks like REST API access is blocked on your website. The REST API access is needed to connect to the Google Sheets. Please enable the REST API access. <a href="https://wppool.dev/docs/possible-errors-and-solutions/" target="_blank">Click</a> here to learn more about this error and the solutions</p> `).setWidth(550).setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Stock Sync with Google Sheet for WooCommerce");
}

function ssgsw_headers() {
    let header = SpreadsheetApp.getActive().getSheetByName(sheetTab).getRange("A1:Z1").getValues();
    header = header[0].filter((column) => column.length);
    return header;
}

function ssgsw_columns($key = null, $reversed = false) {
    let columns = {
        ID: "ID",
        Type: "type",
        Name: "name",
        Stock: "stock",
        "Product URL": "ssgsw_product_url",
        SKU: "sku",
        "Regular price": "regular_price",
        "Sale price": "sale_price",
        "Short description": "post_excerpt",
        "Long description": "post_content",
        Categories: "category",
        "No of Sales": "sales",
        Attributes: "attributes",
    };
    if ($key) {
        if (!$reversed) {
            return columns[$key];
        } else {
            let reverse = {};
            for (let key in columns) {
                reverse[columns[key]] = key;
            }
            return reverse[$key];
        }
    }
    return columns;
}

function ssgsw_available_columns() {
    let columns = ssgsw_columns();
    let headers = ssgsw_headers();
    let keys = {};
    headers.forEach((header) => {
        if (Object.keys(columns).includes(header)) {
            keys[header] = columns[header];
        } else {
            keys[header] = header;
        }
    });
    return keys;
}

function ssgsw_max_column() {
    let maxColumn = SpreadsheetApp.getActive().getSheetByName(sheetTab).getLastColumn();
    return maxColumn;
}

function ssgsw_column_char(index = 0) {
    const alphabet = "abcdefghijklmnopqrstuvwxyz".toUpperCase().split("");
    return alphabet[index - 1] || null;
}

function ssgsw_current_row() {
    let currentCell = SpreadsheetApp.getActive().getSheetByName(sheetTab).getCurrentCell().getA1Notation();
    let row = currentCell.replace(/[^0-9]/g, "");
    return row;
}

function ssgsw_current_column() {
    let currentCell = SpreadsheetApp.getActive().getSheetByName(sheetTab).getCurrentCell().getA1Notation();
    let rowNotation = currentCell.replace(/[0-9]/g, "");
    rowNotation = "abcdefghijklmnopqrstuvwxyz".toUpperCase().split("").indexOf(rowNotation);
    let column = Object.values(ssgsw_available_columns())[rowNotation];
    return column;
}

function ssgsw_column_index(columns) {
    let indexes = [];
    let available_columns = ssgsw_available_columns();
    columns.forEach((column) => {
        let index = Object.values(available_columns).indexOf(column);
        if (index >= 0) index++;
        indexes.push(index);
    });
    return indexes;
}

function ssgsw_format(data) {
    const deletables = ["type", "sales", "category"];
    const keys = ssgsw_ordered_keys();
    data = data.map((row) => {
        return Object.assign.apply({}, keys.map((v, i) => ({
            [v]: row[i]
        })));
    }).map((row) => {
        deletables.forEach((key) => {
            if (key in row) delete row[key];
        });
        return row;
    });
    return data;
}

function ssgsw_get_all_data() {
    var values = SpreadsheetApp.getActive().getSheetByName(sheetTab).getDataRange().getValues();
    values.shift();
    return ssgsw_format(values);
}

function merge_and_parse_image_urls(get_value, formulas) {
    var mergedArray = [];
    for (var i = 0; i < get_value.length; i++) {
        var row = [];
        for (var j = 0; j < get_value[i].length; j++) {
            var cellValue = get_value[i][j];
            var formula = formulas[i][j];
            var imageUrl = ssgsw_extract_image_url_from_formula(formula);
            row.push(imageUrl ? imageUrl : cellValue);
        }
        mergedArray.push(row);
    }
    return mergedArray;
}

function ssgsw_extract_image_url_from_formula(formula) {
    var regex = /=image\("(.*)"/i;
    var matches = formula.match(regex);
    return matches ? matches[1] : null;
}

function ssgsw_get_edited_data(e) {
    var sheet = e.source.getSheetByName(sheetTab);
    var rowStart = e.range.rowStart;
    var rowEnd = e.range.rowEnd;
    var get_all_range = sheet.getRange(rowStart, 1, rowEnd - rowStart + 1, sheet.getLastColumn());
    var get_values = get_all_range.getValues();
    var get_formula = get_all_range.getFormulas();
    var data = merge_and_parse_image_urls(get_values, get_formula);
    return ssgsw_format2(data, rowStart);
}

function ssgsw_format3(data) {
    const deletabless = ["type", "sales", "category"];
    const keyss = ssgsw_ordered_keys();
    keyss.unshift('index_number');
    return data.map((row, index) => {
        let formattedRow = {};
        keyss.forEach((key, i) => {
            formattedRow[key] = row[i];
        });
        deletabless.forEach((key) => {
            if (key in formattedRow) delete formattedRow[key];
        });
        return formattedRow;
    });
}

function ssgsw_format2(data, rowStart) {
    const deletabless = ["type", "sales", "category"];
    const keyss = ssgsw_ordered_keys();
    return data.map((row, index) => {
        let formattedRow = {};
        formattedRow["index_number"] = rowStart + index;
        keyss.forEach((key, i) => {
            formattedRow[key] = row[i];
        });
        deletabless.forEach((key) => {
            if (key in formattedRow) delete formattedRow[key];
        });
        return formattedRow;
    });
}

function ssgsw_ordered_keys() {
    let orderedKeys = [];
    ssgsw_headers().forEach((header) => {
        orderedKeys.push(ssgsw_available_columns()[header]);
    });
    return orderedKeys;
}

function ssgsw_sync_all() {
    let data = ssgsw_get_all_data();
    ssgsw_sync_data(data, 'Products synced successfully', '', 'sync_all');
}

function ssgsw_toast(message = null, title = null) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}

function ssgsw_toast2(message = null, title = null) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, -1);
}

function ssgsw_sync_data(data, message = "Products synced successfully", event = '', sync_all = '') {
    let products = data.filter((row) => {
        return row.name !== '';
    });
    var hasEmptyId = products.some(function(item) {
        return item.ID == '';
    });
    var someEmptyName = data.some(function(item) {
        return item.name == '';
    });
    if (hasEmptyId) {
        if (message != 'Product created successfully!') {
            message = 'Product create and ' + message;
        }
    }
    if (!products.length) {
        ssgsw_toast('Product name cannot be empty! If you want to create a new product', "Warning!");
        ssgsw_fetch_from_WordPress('manual');
        return;
    }
    const arrayLength = products.length;
    ssgsw_toast2("Updating " + arrayLength + " products... Please wait!", 'Loading...');
    let response = UrlFetchApp.fetch(baseURL + "/wp-json/ssgsw/v1/update", {
        method: "POST",
        payload: JSON.stringify({
            products,
            message,
            sync_all,
            arrayLength
        }),
        contentType: "application/json",
        muteHttpExceptions: true,
        headers: {
            SSGSWKEY: "Bearer " + accessToken,
        }
    });
    if (!sync_all) {
        if (response.getResponseCode() == 200) {
            response = JSON.parse(response.getContentText());
            if (response.success) {
                if (response.message == "You couldn't create a new product because the Add new products from Google Sheet feature is not enabled in your settings") {
                    ssgsw_toast(response.message, "Warning!");
                    ssgsw_fetch_from_WordPress('manual');
                } else {
                    ssgsw_toast(response.message, "Success!");
                }
                if (someEmptyName) {
                    ssgsw_toast('Product name cannot be empty! If you want to create a new product', "Warning!");
                    ssgsw_fetch_from_WordPress('manual');
                }
            } else if (response.message) {
                ssgsw_toast(response.message, "Ops error!");
                ssgsw_fetch_from_WordPress('manual');
            }
        } else {
            if (response.getResponseCode() == 401 || response.getResponseCode() == 403) {
                ssgsw_api_access_block_notice();
            } else {
                ssgsw_toast("Something went wrong, please try again", "Oopss!");
            }
        }
    } else {
        ssgsw_toast('Products sync Succesfully!', "Success!");
    }
}

function ssgsw_fetch_from_WordPress(message) {
    let response = UrlFetchApp.fetch(baseURL + "/wp-json/ssgsw/v1/action/?action=sync", {
        method: "GET",
        contentType: "application/json",
        muteHttpExceptions: true,
        headers: {
            SSGSWKEY: "Bearer " + accessToken,
        }
    });
    if (response.getResponseCode() == 200) {
        response = JSON.parse(response.getContentText());
        if (response.success) {
            if (!message) {
                ssgsw_toast("Products fetched from WordPress", "Success!");
            }
        } else if (response.message) {
            ssgsw_toast(response.message, "Ops!");
        }
    } else {
        if (response.getResponseCode() == 401 || response.getResponseCode() == 403) {
            ssgsw_api_access_block_notice();
        } else {
            ssgsw_toast("Something went wrong, please try again", "Oopss!");
        }
    }
}