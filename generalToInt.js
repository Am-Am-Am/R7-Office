//Превращение из общего типа данных в числовой
(function()
{
    let oWorksheet = Api.GetActiveSheet();// Получаем текущий лист
    let test = oWorksheet.Selection; // Объявляем переменную = ее к отбору по ячейке
    test.ForEach(x => { // Создаем Цикл 
        let value = x.GetValue(); // Объявляем переменную = ее к замене 
        if(value === null || value === "" || !Number(value)){ 
            return;
        }
        else{
       
            value = Number(value); // переменная приравнивается к числовому значению
            x.SetValue(value); // выводим переменную
            x.SetNumberFormat("0.00"); // изменяем формат ячейки на числовой и форматируем под определенный вид числа 
        }
    });
})();
