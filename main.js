/**
 * @OnlyCurrentDoc
 */
function createNewCharacterCells () {

    var sDst = SpreadsheetApp.getActiveSheet(), 
        rDst = null;
    var rSrc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reference').getRange("A1:S9");
    var arr = getAllIndexes(sDst.getRange("$A:$A").getValues(), "#");
    var A1Notation = null;

    if ( arr.length == 0 )
        A1Notation = "A1";
    else
        A1Notation= "A"+(arr[arr.length - 1]+sheetConfig.characterSlot.size.rows);

    rDst = sDst.getRange(A1Notation);
    rSrc.copyTo(rDst);
}

var sheetConfig = {
    characterSlot: {
        size: {
            rows : 9,
            columns : 17 
        }
    },
    font : {
        header: {
            size: 18,
            family: "DFKai-SB",
            weight: "bold"
        },
        content: {
            size: 16,
            family: "Microsoft JhengHei",
            weight: "bold"
        }
    },

    backgroundColor : {
        weapon : {
            header : "#e06666",
            content: "#ea9999"
        },

        armor : {
            header : "#f6b26b",
            content: "#f9cb9c"
        },

        sidearmor : {
            header : "#ffd966",
            content: "ffe599"
        },

        accessories : {
            header : "#93c47d",
            content: "#b6d7a8"
        },

        orb : {
            header: "#76a5af",
            content: "#a2c4c9"
        },

        treasure : {
            header : "#6d9eeb",
            content: "#a4c2f4"
        },

        summary : {
            header : "#cc4125",
            content: "#dd7e6b"
        }
    },

}

/* GearOption & SheetWriter. */
var GearOption = function(sheet, startRow) {
    /* Variables */
    // constant
    this.column = 1;
    this.numRows = 6;
    this.numColumns = 2;

    // saves sheet and row index.
    this.sheet = sheet;
    this.row = startRow;
    // set up ranges.
    this.ranges = {};
  
    this.ranges.weapon      = sheet.getRange( this.row + 3, this.column +  1, this.numRows, this.numColumns);
    this.ranges.armor       = sheet.getRange( this.row + 3, this.column +  3, this.numRows, this.numColumns);
    this.ranges.sidearmor   = sheet.getRange( this.row + 3, this.column +  5, this.numRows, this.numColumns);
    this.ranges.accessories = sheet.getRange( this.row + 3, this.column +  7, this.numRows, this.numColumns);
    this.ranges.orb         = sheet.getRange( this.row + 3, this.column +  9, this.numRows, this.numColumns);
    this.ranges.treasure    = sheet.getRange( this.row + 3, this.column + 11, this.numRows, this.numColumns);
    

    // sheet writer
    this.sheetWriter = new SheetWriter( sheet, startRow );

    /* Methods */
    this.startProcess = function () {
        var arrGearOption = this._calcGearOption();
        this.sheetWriter.writeToSheet(arrGearOption);
    };

    this._calcGearOption = function() {
        var valWeapon      = removeEmpty( this.ranges.weapon.getValues() ),
            valArmor       = removeEmpty( this.ranges.armor.getValues() ),
            valSideArmor   = removeEmpty( this.ranges.sidearmor.getValues() ),
            valAccessories = removeEmpty( this.ranges.accessories.getValues() ),
            valOrb         = removeEmpty( this.ranges.orb.getValues() ),
            valTreasures   = removeEmpty( this.ranges.treasure.getValues() );
        var valSum = valWeapon.concat(valArmor, valSideArmor, valAccessories, valOrb, valTreasures);
        var nodupsSum = removeDups(valSum).sort(); 
        return nodupsSum;
    };
};

var SheetWriter = function ( sheet, startRow ){
    /* variables. */
    this.column = 14;
    
    this.sheet = sheet;
    this.startRow = startRow;
  
    /* methods. */
    this.writeToSheet = function( arrGearOption ) {
        var sheet = this.sheet,
            row = this.startRow + 1,
            column = this.column;
        var columnNeeded = 2 * Math.ceil(arrGearOption.length / 6); // each gear option will need two columns.

        this._cellResize( columnNeeded ); // resize cell.
        this._fontCheck(); // check font.

        // write
        arrGearOption.forEach(function(element, index) {
            var type = element[0],
                value = element[1];
            var columnOffset = parseInt(index / 6),
                rowOffset = parseInt(index % 6);

            var targetRange = sheet.getRange( row + rowOffset, column + (2 * columnOffset), 1, 2);

            if (isPercentage(type))
                targetRange.setValues([element]).getCell(1, 2).setNumberFormats([["###.#%"]]);
            else    
                targetRange.setValues([element]).getCell(1, 2).setNumberFormats([["####"]]);
            
        });
    };

    this._cellResize = function( columnNeeded ) {
        this.header = {
            range: this.sheet.getRange(this.startRow, 14).getMergedRanges()[0],
            numColumns: this.sheet.getRange(this.startRow, 14).getMergedRanges()[0].getNumColumns()
        }

        this.content = {
            range: this.sheet.getRange(this.startRow+1, 14, 6, this.header.numColumns),
            numColumns: this.sheet.getRange(this.startRow+1, 14, 6, this.header.numColumns).getNumColumns()
        }

        columnNeeded = columnNeeded < 4 ? 4 : columnNeeded;

        // if current columns number are equal to columns that needed, no need to adjust.
        if ( columnNeeded == this.header.numColumns) return;

        this._resize( columnNeeded );
    };

    this._resize = function ( targetColumns ) {
        var sheet = this.sheet,
            startRow = this.startRow;

        var newHeaderRange = sheet.getRange(startRow, 14, 1, targetColumns);
        var newContentRange = sheet.getRange(startRow+1, 14, 6, targetColumns);

        //header break and merge new range, also save new range into object.
        this.header.range.setBackground(null).breakApart().setBorder(false , null, false , false , false , false );
        newHeaderRange.merge();
        this.header.range = newHeaderRange;

        //content
        this.content.range.setBackground(null).setBorder(false , null, false , false , false , false );
        this.content.range = newContentRange;
    };

    this._fontCheck = function () {
        var conf = sheetConfig;
        this.header.range.setBackground(conf.backgroundColor.summary.header).setBorder(true , null, true , true , true , true ).setFontFamily(conf.font.header.family).setFontSize(conf.font.header.size).setFontWeight(conf.font.header.weight);
        this.content.range.setBackground(conf.backgroundColor.summary.content).setBorder(true , null, true , true , true , true ).setFontFamily(conf.font.content.family).setFontSize(conf.font.content.size).setFontWeight(conf.font.content.weight);
    }
}
/* ./GearOption & SheetWriter */

function startProcess(inputSheet) {
    var sheet = inputSheet ? inputSheet : SpreadsheetApp.getActiveSheet();
    var arrayColumn_A = sheet.getRange("$A:$A").getValues();
    var arrayIndexes = getAllIndexes(arrayColumn_A, "#");

    arrayIndexes.forEach(function(numRow) {
        var gearOption = new GearOption( sheet, numRow); // create this character's gear option object.
        cleanupRange(sheet, numRow); // clean up first.
        gearOption.startProcess(); // start calculate gear option & write result to sheet.
    });
}

function startProcess_all() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function(sheet) {
        var sheetName = sheet.getName();
        startProcess(sheet);
    });
}

function cleanProcess(inputSheet) {
    var sheet = inputSheet ? inputSheet : SpreadsheetApp.getActiveSheet();
    var arrayColumn_A = sheet.getRange("$A:$A").getValues();
    var arrayIndexes = getAllIndexes(arrayColumn_A, "#");
    arrayIndexes.forEach(function(numRow) {
        cleanupRange(sheet, numRow);
    });
}

function cleanProcess_all() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function(sheet) {
        var sheetName = sheet.getName();
        cleanProcess(sheet);
    });

}

function cleanupRange(sheet, startRow) {
    var row = startRow;
    var mergedRange = sheet.getRange(row, 14).getMergedRanges()[0]; // get merged ranges.
    var cleanRange = sheet.getRange(row + 1, 14, 6, mergedRange.getLastColumn() - 14 + 1);

    // clear gear option string.
    cleanRange.clearContent();
    sheet.getRange(row + 1, 15, 6).clearDataValidations();
    sheet.getRange(row + 1, 17, 6).clearDataValidations();
}


/* onOpen, setup menu for Google sheet. */
function onOpen() {
    // Try New Google Sheets method
    try {
        var ui = SpreadsheetApp.getUi();
        ui.createMenu('王之逆襲')
            .addItem("New Character", "createNewCharacterCells")
            .addSeparator()
            .addSubMenu(ui.createMenu('開始計算')
                .addItem("僅當前分頁", 'startProcess')
                .addItem("全部分頁", "startProcess_all")
            )
            .addSeparator()
            .addSubMenu(ui.createMenu('清除結果')
                .addItem("僅當前分頁", "cleanProcess")
                .addItem("全部分頁", "cleanProcess_all")
            )
            .addToUi();
    }

    // Log the error
    catch (e) { Logger.log(e) }

}

/* 
    Utilities functions below.
*/
function removeEmpty(array) {
    var outArray = [];
    array.forEach(function(element) {
        var type = element[0],
            value = element[1];
         if ( type && value && !isNaN(value))
            outArray.push(element);
    });

    return outArray;
}

function removeDups(array) {
    var dict = {};
    var outArray = [];
 
    array.forEach(function(element) {
        var type = element[0],
            value = element[1];
        if (type.length > 0) {
            if ( dict[type] )
                dict[type] += value;
            else
                dict[type] = value;
        }
    });

    Object.keys(dict).forEach(function(type) {
        outArray.push([type, dict[type]]);
    });

    return outArray;
}

function getAllIndexes(arr, val) {
    var indexes = [],
        i;
    for (i = 0; i < arr.length; i++)
        if (arr[i][0] === val)
            indexes.push(i+1);
    return indexes;
}

function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function isPercentage(GearOption) {
    switch (GearOption) {
        case "攻擊":
        case "暴擊傷害量":
        case "防禦":
        case "魔法防禦":
        case "物理防禦":
        case "最大生命":
        case "回復":
        case "每秒魔力回復":
        case "受傷/魔法回復量":
            return true;
        default:
            return false;
    }
}