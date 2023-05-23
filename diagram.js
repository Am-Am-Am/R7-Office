//Создание диаграмм в текстовом редакторе
(function()
{
    let oDocument = Api.GetDocument();
    let oParagraph = oDocument.GetElement(0);
    
    let sType = "bar3D" //Тип диаграммы / "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "stock" | "area" | "areaStacked" | "areaStackedPercent"
    
    let aSeries = [[200, 240, 280],[250, 260, 280]] //Массив данных
    let aSeriesNames = ["Projected Revenue", "Estimated Costs"]//Массив имён данных
    let aCatNames = [2014, 2015, 2016] //Массив имён категорий
    let width = 4051300 // Ширина 
    let height = 2347595 // Длина
    let styleIndex = 24 // Индекс стиля  диаграммы по спецификации OOXML(1 - 48)
    let aNumFormats = ["0", "0.00"]
    
    let oDrawing = Api.CreateChart(sType, aSeries, aSeriesNames, aCatNames, width, height, styleIndex, aNumFormats);
    oDrawing.SetShowPointDataLabel(1, 1, false, false, true, false); // Создание объекта для отрисовки / Индекс значения из массива, над которым будет значение - int/ Индекс столбца, над которым будет значение - int/ Демонстрация имён таблицы - bool/ Демонстрация строк таблицы - bool/ Демонстрация значения данных диаграммы - bool/ Демонстрация процента значений данных - bool
    oParagraph.AddDrawing(oDrawing);
})();
