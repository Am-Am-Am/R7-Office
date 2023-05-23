//Разъединение выбранного диапазона клеток
(function()
{
    Api.GetActiveSheet().GetRange("C3:D10").UnMerge();
})();
