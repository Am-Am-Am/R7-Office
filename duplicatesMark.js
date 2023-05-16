(function () 
{
    let whiteFill = Api.CreateColorFromRGB(255, 255, 255);
    let uniqueColorIndex = 0; // Текущий индекс в цветовом диапазоне
    
    let uniqueColors = [Api.CreateColorFromRGB(255, 255, 0),
        Api.CreateColorFromRGB(204, 204, 255),
        Api.CreateColorFromRGB(0, 255, 0),
        Api.CreateColorFromRGB(0, 128, 128),
        Api.CreateColorFromRGB(192, 192, 192),
        Api.CreateColorFromRGB(255, 204, 0)]; // Массив с цветами

    function getColor() { // Функция, получающая цвета дубликатов 
        if (uniqueColorIndex === uniqueColors.length) {
            uniqueColorIndex = 0;a
        }
        return uniqueColors[uniqueColorIndex++];
    }

    let activeSheet = Api.ActiveSheet; // Получаем текущий лист
    let selection = activeSheet.Selection; // Получаем выделенную область
    let mapValues = {}; // Создаем пустой ассоциативный массив. В нем будет хранится информация о дубликатах.
    let arrRanges = []; //Массив всех клеток
    selection.ForEach(function (range) {
       
        let value = range.GetValue(); // Получаем значение из клеток
        if (!mapValues.hasOwnProperty(value)) {
            mapValues[value] = 0;
        }
        mapValues[value] += 1;
        arrRanges.push(range);
    });
    let value;
    let mapColors = {};
    //Окрашиваем дубликаты
    for (let i = 0; i < arrRanges.length; ++i) {
        value = arrRanges[i].GetValue();
        if (mapValues[value] > 1) {
            if (!mapColors.hasOwnProperty(value)) {
                mapColors[value] = getColor();
            }
            arrRanges[i].SetFillColor(mapColors[value]);
        } else {
            arrRanges[i].SetFillColor(whiteFill);
        }
    }
})();