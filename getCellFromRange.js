let oWorksheet = Api.GetActiveSheet(); //Получение текущего листа
let oRange = oWorksheet.GetRange("A1:C3"); //Получение диапазона ячеек
oRange.GetCells(2, 1).SetFillColor(Api.CreateColorFromRGB(255, 224, 204)); //Получение ячеек из диапазона
