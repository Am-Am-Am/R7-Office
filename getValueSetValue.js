//Получение значения из выбранной области и вставка в нужную ячейку
(function()
{
    let oWorksheet = Api.GetActiveSheet();// Получаем текущий лист
    let test = oWorksheet.Selection; // Объявляем переменную = ее к отбору по ячейке
    test.ForEach(x => {
        let value = x.GetValue();
        
        oWorksheet.GetRange("M1").SetValue(value);
    })
        
})();