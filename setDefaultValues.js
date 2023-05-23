// Очистка значений
let area = secondWorksheet.GetRange("A21:F150"); // Сброс предыдущих значений

function aroundWhiteBorder(el) {
    el.SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Left", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Right", "Medium", Api.CreateColorFromRGB(255, 255, 255));
}

area.ForEach((x) => {
    x.SetValue("");
    aroundWhiteBorder(x);
});
