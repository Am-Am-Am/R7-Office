let selection = Api.GetActiveSheet().Selection; // Помещаем выбранный фрагмент в переменную
let rowsArr = []; // Создаём массив строк
let startRow = selection.GetRow(); // Получаем стартовую строку

// Получаем номер строки каждого выделенного элемента
selection.ForEach((x) => {
    endRow = x.GetRow(); 
});

//Получаем массив с номерами всех выбранных строк
for (startRow; startRow < endRow + 2; startRow++) {
    rowsArr.push(startRow);
}
