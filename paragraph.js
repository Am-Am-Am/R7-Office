//Работа с параграфом в текстовом редакторе
(function()
{
    let oDocument = Api.GetDocument(); //Подключаемся к документу
    let oParagraph = Api.CreateParagraph(); //Создаем параграф
    for(let i = 0; i< 100; i++){ //Цикл
        oParagraph.AddText(`${i}`); //Добавляем с помощью AddText форматированною строку в параграф
        oDocument.Push(oParagraph); //Добавляем параграф на страницу
    }
})();