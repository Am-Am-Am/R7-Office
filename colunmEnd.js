//Перемещение в конец выбранного столбца
(function () 
{
    let activeSheet = Api.ActiveSheet; // Получение текущего листа
    let indexRowMin = 0; // Минимальный индекс строки
    let indexRowMax = 1048576; // Максимальный индекс строки

    let indexCol = 0; // Индекс нужного столбца

    let indexRow = indexRowMax;
    for (; indexRow >= indexRowMin; --indexRow) {
        let range = activeSheet.GetRangeByNumber(indexRow, indexCol);
        if (range.GetValue() && indexRow !== indexRowMax) {
            range = activeSheet.GetRangeByNumber(indexRow + 1, indexCol);
            range.Select();
            break;
        }
    }
})();
