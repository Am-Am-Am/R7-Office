//Функция, получающая ячейку и отрисовывающая границу
function aroundWhiteBorder(el) {
    el.SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Left", "Medium", Api.CreateColorFromRGB(255, 255, 255));
    el.SetBorders("Right", "Medium", Api.CreateColorFromRGB(255, 255, 255));
}
    