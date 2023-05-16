// Функция для удаления границ ячейки - перекрашиваем в белый
let area = secondWorksheet.GetRange("A21:F150"); // Сброс предыдущих значений

area.ForEach((x) => {
x.SetValue("");
aroundWhiteBorder(x);
});
