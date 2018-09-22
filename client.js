/* onOpen, setup menu for Google sheet. */
function onOpen() {
    // Try New Google Sheets method
    try {
        var ui = SpreadsheetApp.getUi();
        ui.createMenu('王之逆襲')
            .addItem("建立新角色欄位", "KingsRaidGearOptionCalculator.createNewCharacterCells")
            .addSeparator()
            .addSubMenu(ui.createMenu('開始計算')
                .addItem("僅當前分頁", 'KingsRaidGearOptionCalculator.startProcess')
                .addItem("全部分頁", "KingsRaidGearOptionCalculator.startProcess_all")
            )
            .addSeparator()
            .addSubMenu(ui.createMenu('清除結果')
                .addItem("僅當前分頁", "KingsRaidGearOptionCalculator.cleanProcess")
                .addItem("全部分頁", "KingsRaidGearOptionCalculator.cleanProcess_all")
            )
            .addToUi();
    }

    // Log the error
    catch (e) { Logger.log(e) }

}