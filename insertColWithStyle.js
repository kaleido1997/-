spread.fromJSON(JSON.parse(sessionStorage.getItem('template')), jsonOptions)

/// begin

// Ref: https://www.grapecity.com.cn/blogs/customize-right-click-insertion
// 定义插入复制样式insertRows/ColsCopyStyle命令：
// Command中使用Transaction实现Undo和Redo，execute时先调用原有"gc.spread.contextMenu.insertRows/insertCols"命令插入行列，
// 然后使用copyTo复制插入行列前的样式。
// copyTo API: http://help.grapecity.com/spread/SpreadSheets11/webframe.html#SpreadJS~GC.Spread.Sheets.Worksheet~copyTo.html
// Or https://demo.grapecity.com.cn/spreadjs/help/api/GC.Spread.Sheets.html#.CopyToOptions
var insertRowsCopyStyle = {
    canUndo: true,
    name: "insertRowsCopyStyle",
    execute: function (context, options, isUndo) {
        var Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
            Commands.undoTransaction(context, options);
            return true;
        } else {
            Commands.startTransaction(context, options);
            var sheet = context.getSheetFromName(options.sheetName);
            sheet.suspendPaint();
            options.cmd = "gc.spread.contextMenu.insertRows"
            context.commandManager().execute(options);
            options.cmd = "insertRowsCopyStyle";

            var beforeRowCount = 0;
            if (options.selections && options.selections.length) {
                var selections = getSortedRowSelections(options.selections)
                for (var i = 0; i < selections.length; i++) {
                    var selection = selections[i];
                    if (selection.row > 0) {
                        for (var row = selection.row + beforeRowCount; row < selection.row + beforeRowCount + selection.rowCount; row++) {
                            sheet.copyTo(selection.row + beforeRowCount - 1, -1, row, -1, 1, -1, GC.Spread.Sheets.CopyToOptions.style | GC.Spread.Sheets.CopyToOptions.span | GC.Spread.Sheets.CopyToOptions.formula);
                        }
                    }
                    beforeRowCount += selection.rowCount;
                }
            }
            sheet.resumePaint();

            Commands.endTransaction(context, options);
            return true;
        }
    }
};

var insertColsCopyStyle = {
    canUndo: true,
    name: "insertColsCopyStyle",
    execute: function (context, options, isUndo) {
        var Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
            Commands.undoTransaction(context, options);
            return true;
        } else {
            Commands.startTransaction(context, options);
            var sheet = context.getSheetFromName(options.sheetName);
            sheet.suspendPaint();
            options.cmd = "gc.spread.contextMenu.insertColumns"
            context.commandManager().execute(options);
            options.cmd = "insertColsCopyStyle";

            var beforeColCount = 0;
            if (options.selections && options.selections.length) {
                var selections = getSortedColSelections(options.selections)
                for (var i = 0; i < selections.length; i++) {
                    var selection = selections[i];
                    if (selection.col > 0) {
                        for (var col = selection.col + beforeColCount; col < selection.col + beforeColCount + selection.colCount; col++) {
                            sheet.copyTo(-1, selection.col + beforeColCount - 1, -1, col, -1, 1, GC.Spread.Sheets.CopyToOptions.style | GC.Spread.Sheets.CopyToOptions.span | GC.Spread.Sheets.CopyToOptions.formula);
                        }
                    }
                    beforeColCount += selection.colCount;
                }
            }
            sheet.resumePaint();

            Commands.endTransaction(context, options);
            return true;
        }
    }
};

// getSortedRow/ColSelections为对selections按照row/col Index排序的方法：
function getSortedRowSelections(selections) {
    var sortedRanges = selections;
    for (var i = 0; i < sortedRanges.length - 1; i++) {
        for (var j = i + 1; j < sortedRanges.length; j++) {
            if (sortedRanges[i].row > sortedRanges[j].row) {
                var temp = sortedRanges[i];
                sortedRanges[i] = sortedRanges[j];
                sortedRanges[j] = temp;
            }
        }
    }
    return sortedRanges;
}

function getSortedColSelections(selections) {
    var sortedRanges = selections;
    for (var i = 0; i < sortedRanges.length - 1; i++) {
        for (var j = i + 1; j < sortedRanges.length; j++) {
            if (sortedRanges[i].col > sortedRanges[j].col) {
                var temp = sortedRanges[i];
                sortedRanges[i] = sortedRanges[j];
                sortedRanges[j] = temp;
            }
        }
    }
    return sortedRanges;
}

// 注册insertRows/ColsCopyStyle命令：
spread.commandManager().register("insertRowsCopyStyle", insertRowsCopyStyle);
spread.commandManager().register("insertColsCopyStyle", insertColsCopyStyle);

// 替换原有插入命令：
function MyContextMenu() {}
MyContextMenu.prototype = new GC.Spread.Sheets.ContextMenu.ContextMenu(spread);
MyContextMenu.prototype.onOpenMenu = function (menuData, itemsDataForShown, hitInfo, spread) {
    itemsDataForShown.forEach(function (item, index) {
        if (item && item.name === "gc.spread.insertRows") {
            item.command = "insertRowsCopyStyle"
        }
        if (item && item.name === "gc.spread.insertColumns") {
            item.command = "insertColsCopyStyle"
        }
    });
    var sheet = spread.getActiveSheet();
    var selections = sheet.getSelections();
    if(selections[0].row === 1){
        //删除某项，也可以通过遍历itemsDataForShown找到对应项位置
        itemsDataForShown.splice(1, 1)
    };
    if(selections[0].col === 1){
        //删除某项，也可以通过遍历itemsDataForShown找到对应项位置
        itemsDataForShown.splice(1, 1)
    }
};
var contextMenu = new MyContextMenu();
spread.contextMenu = contextMenu;

/// end

spread.refresh()
