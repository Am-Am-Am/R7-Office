//Получение номера строки из выделенной области
(function()
{
    let oWorksheet = Api.GetActiveSheet();// Получаем текущий лист
    let test = oWorksheet.Selection;
    let startRow = test.GetRow();
    let endRow;
    test.ForEach(x => {
        endRow = x.GetRow();
    })
    // Объявляем переменную = ее к отбору по ячейке
    
    oWorksheet.GetRange("D20").SetValue(endRow - startRow);
    
})();