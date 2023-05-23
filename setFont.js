//Установление нужного шрифта для всех элементов страницы
(function()
{
    let oDoc = Api.GetDocument(); // Получаем документ
    let elcount = oDoc.GetElementsCount(); // Получаем количество элементов в документе
    for(let i = 0; i < elcount; i++){ // Перебираем циклом все элементы
        let e = oDoc.GetElement(i); 
        let oTextPr = el.GetTextPr();
        oTextPr.SetFontFamily("Comic Sans MS"); // Прописываем название нужного шрифта
        e.SetTextPr(oTextPr);// Устанавливаем нужный шрифт
    }
})();
