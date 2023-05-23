//Слияние выбранного диапазона клеток
(function()
{
    Api.GetActiveSheet().GetRange("A1:B3").Merge(true);
})();
