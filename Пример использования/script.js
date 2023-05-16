(function()
{
  let secondWorksheet = Api.GetSheet("Приёмный акт"); //Получаем второй лист
  let oWorksheet = Api.GetSheet("Пример реестра"); // Получаем текущий лист
  let selection = oWorksheet.Selection; // Закидываем выбранный фрагмент в переменную
  let startRow = selection.GetRow(); // Получаем стартовую строку
  let endRow; // Создаём переменную под конечную строку
  let rowsArr = []; // Создаём массив строк

  let actNumber; // Номер акта
  let actDate; // Дата приёмочного акта, ячейка B2
  let station; // Станция отправления, ячейка B6
  let storage; // Склад, ячейка B4
  let shipper; // Грузоотправитель, ячейка B8
  let provider; // Поставщик, ячейка B10
  let carriage; // Вагон, ячейка B12
  let documentT; // Наименование, номер товарно транспортного документа, ячейка B16
  let releaseDate; // Дата раскредитации, ячейка B18

  let itemName; // Наименование товаро-материальных ценностей,
  let units; // Единицы измерения,
  let count; // Количество,
  let realCount; // Количество фактическое,
  let comment; // Комментарий

  let deliverer; // Приемосдатчик, ячейка F28
  let manager; // Заведующий складом, ячейка F30
  let engineer; // Инженер

  let area = secondWorksheet.GetRange("A21:F150"); // Сброс предыдущих значений
// Функция для удаления границ ячейки - перекрашиваем в белый
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

  selection.ForEach((x) => {
    endRow = x.GetRow(); // Получаем номер строки каждого выделенного элемента
  });

  function checkArr(arr) {
    // Функция, проверяющая элементы массива на одинаковость и возвращающая true или false
    let first = arr[0];
    return arr.every((el) => el === first);
  }
//Функция для создания границ ячейки - перекрашиваем в чёрный
  function aroundBorder(el) {
    el.SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0, 0, 0));
    el.SetBorders("Top", "Medium", Api.CreateColorFromRGB(0, 0, 0));
    el.SetBorders("Left", "Medium", Api.CreateColorFromRGB(0, 0, 0));
    el.SetBorders("Right", "Medium", Api.CreateColorFromRGB(0, 0, 0));
  }

  //Получаем массив с номерами всех выбранных строк
  for (startRow; startRow < endRow + 2; startRow++) {
    rowsArr.push(startRow);
  }

  let dateArr = []; // Создаём массив дат
  let numberArr = []; // Создаём массив номеров приёмочного акта
  let storageArr = []; //Создаём массив складов
  let stationArr = []; //Создаём массив станций
  let shipperArr = []; //Создаём массив грузоотправителей
  let providerArr = []; //Создаём массив поставщиков
  let carriageArr = []; //Создаём массив вагонов
  let releaseDateArr = []; //Создаём массив дат раскредитации
  let delivererArr = []; //Создаём массив приемосдатчиков
  let engineerArr = []; //Создаём массив инженеров


  // Собираем значения с каждой из позиций и закидываем в одноимённый массив
  for (let i = 0; i < rowsArr.length; i++) {
    dateArr.push(oWorksheet.GetRange(`D${rowsArr[i + 1]}`).GetText());
    numberArr.push(oWorksheet.GetRange(`C${rowsArr[i + 1]}`).GetText());
    storageArr.push(oWorksheet.GetRange(`F${rowsArr[i + 1]}`).GetText());
    stationArr.push(oWorksheet.GetRange(`J${rowsArr[i + 1]}`).GetText());
    shipperArr.push(oWorksheet.GetRange(`K${rowsArr[i + 1]}`).GetText());
    providerArr.push(oWorksheet.GetRange(`L${rowsArr[i + 1]}`).GetText());
    carriageArr.push(oWorksheet.GetRange(`M${rowsArr[i + 1]}`).GetText());
    releaseDateArr.push(oWorksheet.GetRange(`P${rowsArr[i + 1]}`).GetText());
    delivererArr.push(oWorksheet.GetRange(`G${rowsArr[i + 1]}`).GetText());
    engineerArr.push(oWorksheet.GetRange(`I${rowsArr[i + 1]}`).GetText());
    
    //Массив имён категорий
    let names = [
      "Даты ",
      "Номера приемочного акта ",
      "Склады ",
      "Станции ",
      "Грузоотправители ",
      "Поставщики ",
      "Вагоны ",
      "Даты раскредитации ",
      "Приемосдатчики ",
      "Инженеры ",
    ];
    //Создаём массив с массивами категорий
    let dataArr = [
      dateArr,
      numberArr,
      storageArr,
      stationArr,
      shipperArr,
      providerArr,
      carriageArr,
      releaseDateArr,
      delivererArr,
      engineerArr,
    ];
    let resultError = "";
    //Пробегаем по массиву с массивами категорий, для каждого из которых вызываем функцию, проверяющую одинаковость значений. 
    //Если значения не одинаковы, то по индексу неверного массива обращаемся в массив имён категорий. Это для отображения
    for (let dataArrCount = 0; dataArrCount < dataArr.length; dataArrCount++) {
      if (!checkArr(dataArr[dataArrCount])) {
        resultError += names[dataArrCount];
        oWorksheet.GetRange("F4").SetValue("Ошибка: " + resultError);
        secondWorksheet.GetRange("I4").SetValue("Ошибка: " + resultError);
      }
    }

    //По номерам выбранных строк получаем все выбранные номера приэмочного акта
    //Всё основное происходит в этом цикле. Закидываем в функцию проверки одинаковости номеров наш массив
    if (resultError.length == 0) {
      // При успехе создаём нужные переменные
      oWorksheet.GetRange("F2").SetValue("");
      oWorksheet.GetRange("F4").SetValue("");
      secondWorksheet.GetRange("I4").SetValue("");
      oWorksheet.GetRange("F2").SetValue("Значения совпадают");
      secondWorksheet.GetRange("I2").SetValue("Значения совпадают");

      for (let j = 0; j < numberArr.length; j++) {
        //Номер акта
        actNumber = oWorksheet.GetRange(`C${rowsArr[j + 1]}`).GetText();
        //Дата приёмочного акта
        actDate =
          "от г. " + oWorksheet.GetRange(`D${rowsArr[j + 1]}`).GetText();
        //Станция отправления, получение значения
        station = oWorksheet.GetRange(`J${rowsArr[j + 1]}`).GetText();
        //Склад
        storage = oWorksheet.GetRange(`F${rowsArr[j + 1]}`).GetText();
        //Грузоотправитель
        shipper = oWorksheet.GetRange(`K${rowsArr[j + 1]}`).GetText();
        // Поставщик
        provider = oWorksheet.GetRange(`L${rowsArr[j + 1]}`).GetText();
        //Вагон
        carriage = oWorksheet.GetRange(`M${rowsArr[j + 1]}`).GetText();
        //Наименование, номер товарно транспортного документа
        documentT =
          oWorksheet.GetRange(`N${rowsArr[j + 1]}`).GetText() +
          " " +
          oWorksheet.GetRange(`O${rowsArr[j + 1]}`).GetText();
        //Дата раскредитации
        releaseDate = oWorksheet.GetRange(`P${rowsArr[j + 1]}`).GetText();
        //Наименование товаро-материальных ценностей
        itemName = oWorksheet.GetRange(`Q${rowsArr[j + 1]}`).GetText();
        //Единицы измерения
        units = oWorksheet.GetRange(`R${rowsArr[j + 1]}`).GetText();
        //Количество
        count = oWorksheet.GetRange(`S${rowsArr[j + 1]}`).GetText();
        //Количество фактическое
        realCount = oWorksheet.GetRange(`T${rowsArr[j + 1]}`).GetText();
        //Комментарий
        comment = oWorksheet.GetRange(`V${rowsArr[j + 1]}`).GetText();

        //Работа с таблицей элементов. Отрисовка внутри цикла.
        secondWorksheet.GetRange(`A${21 + j}`).SetValue(j + 1); // Номер строки
        secondWorksheet.GetRange(`B${21 + j}`).SetValue(itemName); // Наименование товаро-материальных ценностей
        secondWorksheet.GetRange(`C${21 + j}`).SetValue(units); // Единицы измерения
        secondWorksheet.GetRange(`D${21 + j}`).SetValue(count); // Количество
        secondWorksheet.GetRange(`E${21 + j}`).SetValue(realCount); // Количество фактическое
        secondWorksheet.GetRange(`F${21 + j}`).SetValue(comment); // Комментарий

        deliverer = "";
        manager = "";
        engineer = "";

        //Приемосдатчик
        deliverer += oWorksheet.GetRange(`G${rowsArr[j + 1]}`).GetText();
        //Заведующий складом
        manager += oWorksheet.GetRange(`H${rowsArr[j + 1]}`).GetText();
        //Инженер
        engineer += oWorksheet.GetRange(`I${rowsArr[j + 1]}`).GetText();

        // Формирование нижней части документа, привязка к j нужна для отступа от элементов

        aroundBorder(secondWorksheet.GetRange(`A${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`B${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`C${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`D${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`E${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`F${21 + j}`));
        // Сначала ставим пустые значения
         // Сначала ставим пустые значения
         secondWorksheet.GetRange(`A${28 + j}`).SetValue("");
         secondWorksheet.GetRange(`A${28 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
         secondWorksheet.GetRange(`A${30 + j}`).SetValue("");
         secondWorksheet.GetRange(`A${30 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
         secondWorksheet.GetRange(`A${32 + j}`).SetValue("");
         secondWorksheet.GetRange(`A${32 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
         secondWorksheet.GetRange(`F${29 + j}`).SetValue("");
         secondWorksheet.GetRange(`F${29 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
         secondWorksheet.GetRange(`F${31 + j}`).SetValue("");
         secondWorksheet.GetRange(`F${31 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
         secondWorksheet.GetRange(`F${33 + j}`).SetValue("");
         secondWorksheet.GetRange(`F${33 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
 
         // Потом ставим нужные значения
         secondWorksheet.GetRange(`A${29 + j}`).SetValue("Приемосдатчик");
         secondWorksheet.GetRange(`A${29 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
         secondWorksheet.GetRange(`A${31 + j}`).SetValue("Заведующий складом");
         secondWorksheet.GetRange(`A${31 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
         secondWorksheet.GetRange(`A${33 + j}`).SetValue("Инженер");
         secondWorksheet.GetRange(`A${33 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
         secondWorksheet.GetRange(`F${30 + j}`).SetValue(deliverer);
         secondWorksheet.GetRange(`F${30 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));
         secondWorksheet.GetRange(`F${32 + j}`).SetValue(manager);
         secondWorksheet.GetRange(`F${32 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));
         secondWorksheet.GetRange(`F${34 + j}`).SetValue(engineer);
         secondWorksheet.GetRange(`F${34 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));

        
      }

      secondWorksheet.GetRange("C1").SetValue(actNumber); //Номер акта
      secondWorksheet.GetRange("B2").SetValue(actDate); //Дата приёмочного акта
      secondWorksheet.GetRange("B6").SetValue(station); //Станция отправления
      secondWorksheet.GetRange("B4").SetValue(storage); //Склад
      secondWorksheet.GetRange("B8").SetValue(shipper); //Грузоотправитель
      secondWorksheet.GetRange("B10").SetValue(provider); //Поставщик
      secondWorksheet.GetRange("B12").SetValue(carriage); //Вагон
      secondWorksheet.GetRange("B16").SetValue(documentT); //Наименование, номер товарно транспортного документа
      secondWorksheet.GetRange("B18").SetValue(releaseDate); //Дата раскредитации
    } else {
      let area = secondWorksheet.GetRange("A21:F150"); // Сброс предыдущих значений

      oWorksheet.GetRange("F2").SetValue("");

      area.ForEach((x) => {
        x.SetValue("");
      });

      oWorksheet.GetRange("F2").SetValue("Значения не совпадают"); //Разные номера - пишем об этом около кнопки
      secondWorksheet.GetRange("I2").SetValue("Значения не совпадают");

      actNumber = "";
      actDate = "от г. ";
      station = "";
      storage = "";
      shipper = "";
      provider = "";
      carriage = "";
      documentT = "";
      releaseDate = "";

      count = 0;
      for (let j = 0; j < numberArr.length; j++) {
        //Номер акта
        actNumber += oWorksheet.GetRange(`C${rowsArr[j + 1]}`).GetText() + " ";
        //Дата приёмочного акта
        actDate += oWorksheet.GetRange(`D${rowsArr[j + 1]}`).GetText() + " ";
        //Станция отправления, получение значения
        station += oWorksheet.GetRange(`J${rowsArr[j + 1]}`).GetText() + " ";
        //Склад
        storage += oWorksheet.GetRange(`F${rowsArr[j + 1]}`).GetText() + " ";
        //Грузоотправитель
        shipper += oWorksheet.GetRange(`K${rowsArr[j + 1]}`).GetText() + " ";
        // Поставщик
        provider += oWorksheet.GetRange(`L${rowsArr[j + 1]}`).GetText() + " ";
        //Вагон
        carriage += oWorksheet.GetRange(`M${rowsArr[j + 1]}`).GetText() + " ";
        //Наименование, номер товарно транспортного документа
        documentT +=
          oWorksheet.GetRange(`N${rowsArr[j + 1]}`).GetText() +
          " " +
          oWorksheet.GetRange(`O${rowsArr[j + 1]}`).GetText() +
          " ";
        //Дата раскредитации
        releaseDate +=
          oWorksheet.GetRange(`P${rowsArr[j + 1]}`).GetText() + " ";
        //Наименование товаро-материальных ценностей
        itemName = oWorksheet.GetRange(`Q${rowsArr[j + 1]}`).GetText();
        //Единицы измерения
        units = oWorksheet.GetRange(`R${rowsArr[j + 1]}`).GetText();
        //Количество
        count = oWorksheet.GetRange(`S${rowsArr[j + 1]}`).GetText();
        //Количество фактическое
        realCount = oWorksheet.GetRange(`T${rowsArr[j + 1]}`).GetText();
        //Комментарий
        comment = oWorksheet.GetRange(`V${rowsArr[j + 1]}`).GetText();

        //Работа с таблицей элементов. Отрисовка внутри цикла.
        aroundBorder(secondWorksheet.GetRange(`A${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`B${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`C${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`D${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`E${21 + j}`));
        aroundBorder(secondWorksheet.GetRange(`F${21 + j}`));

        secondWorksheet.GetRange(`A${21 + j}`).SetValue(j + 1); // Номер строки
        secondWorksheet.GetRange(`B${21 + j}`).SetValue(itemName); // Наименование товаро-материальных ценностей
        secondWorksheet.GetRange(`C${21 + j}`).SetValue(units); // Единицы измерения
        secondWorksheet.GetRange(`D${21 + j}`).SetValue(count); // Количество
        secondWorksheet.GetRange(`E${21 + j}`).SetValue(realCount); // Количество фактическое
        secondWorksheet.GetRange(`F${21 + j}`).SetValue(comment); // Комментарий

        deliverer = "";
        manager = "";
        engineer = "";

        //Приемосдатчик
        deliverer += oWorksheet.GetRange(`G${rowsArr[j + 1]}`).GetText();
        //Заведующий складом
        manager += oWorksheet.GetRange(`H${rowsArr[j + 1]}`).GetText();
        //Инженер
        engineer += oWorksheet.GetRange(`I${rowsArr[j + 1]}`).GetText();

        // Формирование нижней части документа, привязка к j нужна для отступа от элементов

        // Сначала ставим пустые значения
        secondWorksheet.GetRange(`A${28 + j}`).SetValue("");
        secondWorksheet.GetRange(`A${28 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
        secondWorksheet.GetRange(`A${30 + j}`).SetValue("");
        secondWorksheet.GetRange(`A${30 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
        secondWorksheet.GetRange(`A${32 + j}`).SetValue("");
        secondWorksheet.GetRange(`A${32 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
        secondWorksheet.GetRange(`F${29 + j}`).SetValue("");
        secondWorksheet.GetRange(`F${29 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
        secondWorksheet.GetRange(`F${31 + j}`).SetValue("");
        secondWorksheet.GetRange(`F${31 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
        secondWorksheet.GetRange(`F${33 + j}`).SetValue("");
        secondWorksheet.GetRange(`F${33 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));

        // Потом ставим нужные значения
        secondWorksheet.GetRange(`A${29 + j}`).SetValue("Приемосдатчик");
        secondWorksheet.GetRange(`A${29 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
        secondWorksheet.GetRange(`A${31 + j}`).SetValue("Заведующий складом");
        secondWorksheet.GetRange(`A${31 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
        secondWorksheet.GetRange(`A${33 + j}`).SetValue("Инженер");
        secondWorksheet.GetRange(`A${33 + j}`).SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(0,0,0));
        secondWorksheet.GetRange(`F${30 + j}`).SetValue(deliverer);
        secondWorksheet.GetRange(`F${30 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));
        secondWorksheet.GetRange(`F${32 + j}`).SetValue(manager);
        secondWorksheet.GetRange(`F${32 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));
        secondWorksheet.GetRange(`F${34 + j}`).SetValue(engineer);
        secondWorksheet.GetRange(`F${34 + j}`).SetBorders("Top", "Medium", Api.CreateColorFromRGB(0,0,0));
      }

      secondWorksheet.GetRange("C1").SetValue(actNumber); //Номер акта
      secondWorksheet.GetRange("B2").SetValue(actDate); //Дата приёмочного акта
      secondWorksheet.GetRange("B6").SetValue(station); //Станция отправления
      secondWorksheet.GetRange("B4").SetValue(storage); //Склад
      secondWorksheet.GetRange("B8").SetValue(shipper); //Грузоотправитель
      secondWorksheet.GetRange("B10").SetValue(provider); //Поставщик
      secondWorksheet.GetRange("B12").SetValue(carriage); //Вагон
      secondWorksheet.GetRange("B16").SetValue(documentT); //Наименование, номер товарно транспортного документа
      secondWorksheet.GetRange("B18").SetValue(releaseDate); //Дата раскредитации
    }
  }
  
})();