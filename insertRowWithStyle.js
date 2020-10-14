spread.fromJSON(JSON.parse(sessionStorage.getItem('template')), jsonOptions)

/// begin

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

spread.commandManager().register("insertRowsCopyStyle", insertRowsCopyStyle);

function MyContextMenu() {}
MyContextMenu.prototype = new GC.Spread.Sheets.ContextMenu.ContextMenu(spread);
MyContextMenu.prototype.onOpenMenu = function (menuData, itemsDataForShown, hitInfo, spread) {
    itemsDataForShown.forEach(function (item, index) {
        if (item && item.name === "gc.spread.insertRows") {
            item.command = "insertRowsCopyStyle"
        }
    });
    var sheet = spread.getActiveSheet();
    var selections = sheet.getSelections();
    if(selections[0].row === 1){
        //删除某项，也可以通过遍历itemsDataForShown找到对应项位置
        itemsDataForShown.splice(1, 1)
    }

};
var contextMenu = new MyContextMenu();
spread.contextMenu = contextMenu;

/// end

spread.refresh()
