//Изменение цвета фона выбранного диапазона
(function()
{
    Api.GetActiveSheet().GetRange("B3").SetFillColor(Api.CreateColorFromRGB(0, 0, 250));
})();