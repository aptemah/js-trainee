

function MyExcel() {

  var self = this;

  this.viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  this.viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

  this.charArray = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

};


MyExcel.prototype.init = function (block, settings) {

  var self = this;

  self.tableName = block;//определяем имя таблицы. Объект в Local Storage будет называться так же
  this.parentBlock = document.getElementById(block);

  if (localStorage.getItem(self.tableName)) {

    self.sheetObject = JSON.parse(localStorage[self.tableName]);

  } else {

    var tableName = "files/" + this.tableName + ".json";
    var xhr = new XMLHttpRequest();
    xhr.open('GET', tableName, false);
    xhr.send();
    if (xhr.status != 200) {

      self.sheetObject = {};

      } else {
        try {
          self.sheetObject = JSON.parse(xhr.responseText);
          localStorage[this.tableName] = xhr.responseText;
        } catch(err) {
          self.sheetObject = {};
        }
    }
  };

  self.tabsBlock = document.createElement("DIV");
  self.tabsBlock.className = "tabs-wrap";
  this.parentBlock.parentNode.insertBefore(self.tabsBlock, this.parentBlock.nextSibling);
  self.tabGenerating();

  if (settings.width != undefined && settings.width != null) { widthOfTable = settings.width} else {
    //из ширины вычитаем ширину скролла
    var widthOfTable = this.viewportWidth - self.scrollWidth();
  }
  if (settings.height != undefined && settings.height != null) { heightOfTable = settings.height} else {
    //из высоты вычитаем отступ сверху и высоту табов
    var heightOfTable = this.viewportHeight - parseInt(window.getComputedStyle(this.parentBlock).getPropertyValue('margin-top')) - self.tabsBlock.clientHeight;
  }

  this.parentBlock.style.width  = (widthOfTable + "px");
  this.parentBlock.style.height = (heightOfTable + "px");
  self.cellWidth = 80;
  self.cellHeight = 20;

  self.creatingTableDom(false, settings.rows, settings.columns);

  this.parentBlock.addEventListener("scroll", function(){

    if (self.parentBlock.scrollTop == self.parentBlock.scrollHeight - self.parentBlock.clientHeight) {

      var cellsQuantity = self.table.rows[0].children.length;

      for (var i = 0; i < 5; i++) {
        var row = self.table.insertRow();
        //добавления в таблицу разметки
        var rowGrid = self.gridTableElementRows.insertRow();
        rowGrid.insertCell();
        self.gridTableElementRows.rows[row.rowIndex].cells[0].textContent = row.rowIndex + 1;

        for (var j = 0; j < cellsQuantity; j++) {
          var cell = row.insertCell();

          //заполняем ячейки, если они присутствуют в sheetObject
          if (self.sheetObject[self.currentSheet][row.rowIndex]) {
            if (self.sheetObject[self.currentSheet][row.rowIndex][cell.cellIndex]) {
              cell.textContent = self.sheetObject[self.currentSheet][row.rowIndex][cell.cellIndex]
            }
          }

        }
      }

    }

    //Горизонтальный скролл
    if (self.parentBlock.scrollLeft == self.parentBlock.scrollWidth - self.parentBlock.clientWidth) {

      var numberOfAddingColumns = 5;
      //добавления в таблицу разметки
      for (var i = 0; i < numberOfAddingColumns; i++) {
        var cell = self.gridTableElementColumns.rows[0].insertCell();
        var cellIndex = parseInt(cell.cellIndex);
        self.gridTableElementColumns.rows[0].cells[cellIndex].textContent = self.indexToChar(cellIndex);
      }
      //добавления в основную таблицу
      for (var row in self.table.rows) {
        if (!self.table.rows.hasOwnProperty(row)) continue;
        for (var i = 0; i < numberOfAddingColumns; i++) {
          var cell = self.table.rows[row].insertCell();

          //заполняем ячейки, если они присутствуют в sheetObject
          if (self.sheetObject[self.currentSheet][row]) {
            if (self.sheetObject[self.currentSheet][row][cell.cellIndex]) {
              cell.textContent = self.sheetObject[self.currentSheet][row][cell.cellIndex]
            }
          }

        }
      }

    }

  });

  this.getCoords();
  this.makeGrid(this.table);

};


MyExcel.prototype.indexToChar = function (index) {

  var self = this;
  var word = "", result;

  this.oneMoreTime = function (lastResult) {

    if (lastResult == undefined) {lastResult = index}
    result = Math.floor(lastResult / 26);
    if (result > 26) {
      var middleIndex = result % 26;
      word = word + this.charArray[middleIndex - 1];
      this.oneMoreTime(result)
    } else {
      word = this.charArray[result - 1] + word;
      if (result > 26) { this.oneMoreTime(result) } else { this.theLastChar(index % 26) };
    }

  }

  this.theLastChar = function (result) {

    if (result == undefined) {result = index} 
    word = word + self.charArray[result];
  }

  if (index < 26) {this.theLastChar()} else {this.oneMoreTime()}

  return word;

};


MyExcel.prototype.charToIndex = function (char) {

  var index = 0, rank = 0, resultArray = [];
  var strArray = char.split("");

  strArray.forEach( (currentChar) => {
    var indexOfCurrentChar = this.charArray.indexOf(currentChar.toUpperCase());
    resultArray.push(indexOfCurrentChar);
  });

  resultArray.forEach( (item, i, arr) => {
    var add = item;
    for (var j = 0; j < arr.length - i - 1; j++) {
      if (j == 0) {add++}
      add = 26 * add
    }
    index = index + add;
  });
  return index;

};


MyExcel.prototype.creatingTableDom = function (currentSheet, rows, columns) {
  //создание DOM таблицы
  if (currentSheet == undefined) {
    this.currentSheet = this.tabsBlock.querySelector(":checked").value;
  } else { this.currentSheet = currentSheet };

  var maxRowNumber  = 0;
  var maxCellNumber = 0;

  if (this.table) {this.parentBlock.removeChild(this.table)}

  if (!rows) {
    var rowsInitialAmount = Math.ceil(this.viewportHeight / this.cellHeight);
    if (rowsInitialAmount < maxRowNumber) {rowsInitialAmount = maxRowNumber + 1};
  } else { rowsInitialAmount = rows }

  if (!columns) {
    var columnsInitialAmount = Math.ceil(this.viewportWidth / this.cellWidth);
    if (columnsInitialAmount < maxCellNumber) {columnsInitialAmount = maxCellNumber + 1};
  } else { columnsInitialAmount = columns }

  this.globalInput = document.createElement("INPUT");
  this.globalInput.id = this.tableName + "-global-input";
  this.globalInput.className = "global-input";
  this.parentBlock.appendChild(this.globalInput);

  var tableFragment = document.createDocumentFragment();
  this.table = document.createElement("TABLE");
  this.table.className = "super-table";
  tableFragment.appendChild(this.table);

  for (var i = 0; i < rowsInitialAmount; i++) {
    var row = this.table.insertRow();
    for (var j = 0; j < columnsInitialAmount; j++) {
      row.insertCell();
    }
  }

  this.parentBlock.appendChild(this.table);

  this.fillCells(this.currentSheet);//Заполняем ячейки
  this.cellSelect();

};


MyExcel.prototype.cellSelect = function () {

  var self = this;

  this.hoverFunction = function (e) {

    if (e.target.tagName == "TD") {

      var rowIndex = e.target.parentElement.rowIndex;
      var cellIndex = e.target.cellIndex;
      var targetCell = self.table.tBodies[0].rows[rowIndex].cells[cellIndex];
      var input = document.createElement("INPUT");

      self.selectedCell = self.indexToChar(cellIndex - 1) + (rowIndex - 1);

      targetCell.className += "hover";

      window.addEventListener("keydown", inputCreating);
      targetCell.addEventListener("click", inputCreating);
      self.table.addEventListener("click", deleteHoverClass);


      function inputCreating (e) {

        self.table.removeEventListener("click", self.hoverFunction);
        targetCell.removeEventListener("click", inputCreating);
        self.table.removeEventListener("click", deleteHoverClass);
        setTimeout(function(){self.table.addEventListener("click", addReference)});

        input.addEventListener("keypress", enterButton);
        input.addEventListener("keyup", escapeButton);
        try {//если кликаешь по пустой ячейке (которой, естественно, нет в sheetObject, сыпятся ошибки. Чтоб не делать проверки, решил заюзать try catch)
          if (self.sheetObject[self.currentSheet][rowIndex][cellIndex]) {input.value = self.sheetObject[self.currentSheet][rowIndex][cellIndex];
          } else {
            input.value = "";
          }
        } catch(err) {}

        self.oldValue = targetCell.textContent;//записываем старое значение, чтобы была возможность отменить ввод по escape
        targetCell.textContent = null;
        targetCell.appendChild(input);
        input.focus();
        inputBinding();
        window.removeEventListener("keydown", inputCreating);

      }


      function inputBinding () {

        var globalInput = self.globalInput;
        globalInput.className = globalInput.className + " active"

        input.oninput = function() {
          globalInput.value = input.value;
        };

        globalInput.oninput = function() {
          input.value = globalInput.value;
        };

      };


      function addReference (e) {

        if (e.target != input) {
          if (input.value.search(/^=/) != "-1") {
            input.value = input.value + self.selectedCellCoords;
            input.focus();
          } else { saveValue() }
        } else { saveValue() }

      };

      function deleteHoverClass() {

        targetCell.className = targetCell.className.replace(/\bhover\b/,'');
        window.removeEventListener("keydown", inputCreating);
        targetCell.removeEventListener("click", inputCreating);
        self.table.removeEventListener("click", deleteHoverClass);

      }


      function enterButton (e) {

        var key = e.which || e.keyCode;
        if (key === 13) {
          saveValue();
        }

      }


      function escapeButton (e) {

        var key = e.which || e.keyCode;
        if (key === 27) {
          input.value = self.oldValue;
          deleteHoverClass();
          input.parentElement.textContent = input.value;
          input = null;

          self.table.removeEventListener("click", addReference);
          self.table.removeEventListener("click", deleteHoverClass);
          self.table.addEventListener("click", self.hoverFunction);
          inputUnfocus();
        }

      }


      function saveValue() {

        deleteHoverClass();

        var currentSheet = self.tabsBlock.querySelector(":checked").value;

        if (!self.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
          self.sheetObject[currentSheet][rowIndex] = {};
        }

        self.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;
        self.updateSheet();

        if (self.ifFormula(input.value)) {
          input.parentElement.textContent = self.formulaParse( input.value );
          self.table.dispatchEvent(widgetEvent);

          //добавляем ячейки со ссылками в специальный объект
          var formulaArray = input.value.match(/[a-z]{1,}\d{1,}/gi);
          console.log(formulaArray);
          if (formulaArray != null)
          formulaArray.forEach(function(i){
          var char = i.match(/[a-z]{1,}/i)[0],
              digit = i.match(/\d{1,}/i)[0];
            self.linkCells[parseInt( digit )] = self.charToIndex( char );
          });

          //добавляем ячейки с формулами в специальный объект
          self.formulaCells[targetCell.parentElement.rowIndex] = targetCell.cellIndex;

        } else {
          input.parentElement.textContent = input.value
        }

        //Триггерим пересчет формул, если мы изменяем ячейку, на которую ссылаемся из других ячеек (self.linkCells)
        if (self.linkCells.hasOwnProperty(targetCell.parentElement.rowIndex + 1)) {

          if (self.linkCells[targetCell.parentElement.rowIndex + 1] == targetCell.cellIndex) {console.log(self.linkCells);self.table.dispatchEvent(widgetEvent)}
        }

        input = null;

        self.table.removeEventListener("click", addReference);
        self.table.removeEventListener("click", deleteHoverClass);
        self.table.addEventListener("click", self.hoverFunction);
        inputUnfocus();
      };


      function inputUnfocus () {

        self.globalInput.className = self.globalInput.className.replace(/\b active\b/,'');

      };

    }
  };

  this.table.addEventListener("click", this.hoverFunction);

};


MyExcel.prototype.getCoords = function () {

  var self = this;

  this.table.addEventListener("click", (e) => {

    if (e.target.tagName == "TD") {
      var rowIndex = e.target.parentElement.rowIndex;
      var cellIndex = e.target.cellIndex;

      self.selectedCellCoords = self.indexToChar( cellIndex ) + (rowIndex + 1);

    }

  });

};


MyExcel.prototype.makeGrid = function () {

  var self = this;

  if (this.gridTableElementRows) {
    self.parentBlock.removeChild(this.gridTableElementRows);
  }
  if (this.gridTableElementColumns) {
    self.parentBlock.removeChild(this.gridTableElementColumns);
  }

  this.addNewRow = function (numberOfRow) {

    this.gridTableElementRows.rows[numberOfRow].cells[0].textContent = numberOfRow + 1;

  }

  this.addNewColumn = function (numberOfColumn) {

    this.gridTableElementColumns.rows[0].cells[numberOfColumn].textContent = this.indexToChar(numberOfColumn);

  }

  var tableRows = Array.prototype.slice.call(this.table.rows);
  var firstRowCells = Array.prototype.slice.call(tableRows[0].cells);

  //таблица для строк
  this.gridTableElementRows = document.createElement("TABLE");
  this.gridTableElementRows.className = "super-table-rows";
  this.parentBlock.appendChild(this.gridTableElementRows);

  for (var i = 0; i < tableRows.length; i++) {
    var row = this.gridTableElementRows.insertRow();
    row.insertCell();
    this.addNewRow(i);
  }

  //таблица для столбцов
  this.gridTableElementColumns = document.createElement("TABLE");
  this.gridTableElementColumns.className = "super-table-columns";
  this.parentBlock.appendChild(this.gridTableElementColumns);

  var row = this.gridTableElementColumns.insertRow();

  for (var i = 0; i < firstRowCells.length; i++) {
    row.insertCell();
    this.addNewColumn(i);
  }

  this.parentBlock.addEventListener("scroll", function(){

    self.gridTableElementRows.style.left = self.parentBlock.scrollLeft + "px";
    self.gridTableElementColumns.style.top = self.parentBlock.scrollTop + "px";

  });

};


MyExcel.prototype.tabDelete = function (index, elements) {

  //удаляем таб
  var tabsWrap = this.tabsBlock;
  elements.forEach((item, i, arr) => {
    tabsWrap.removeChild(item);
  });

  //удаляем объект листа и обновляем localStorage
  this.deleteSheet(index);

  //устанавливаем нужный таб и перерисовываем страницу
  var sheetToRedirect = this.setReaddressSheet(index);

  if (sheetToRedirect == undefined) {// если undefined — значит это был единственный лист, и нужно создать новый

    this.tabCreate(0, "sheet0", true);
    this.sheetCreate(0);
    this.creatingTableDom();

  } else {

    tabsWrap.querySelector("#" + this.tableName + "tab" + sheetToRedirect).checked = true;
    this.creatingTableDom();
    this.setCurrentSheet(sheetToRedirect);

  }

};


MyExcel.prototype.setReaddressSheet = function (index) {

  //выясняем на какой таб будет переадресация после удаления
  var previousElement = 0;//кодовое число, будет использоваться в случае если это первый лист
  for (var sheet in this.sheetObject) {

    if (!this.sheetObject.hasOwnProperty(sheet)) continue;
    if (previousElement != undefined) {
      if (sheet != index) {//если первый лист, проверяем не единственный ли это лист
        return sheet;
        break;
      }
      return previousElement;
      break;
    }
    previousElement = sheet;
  }

};


MyExcel.prototype.tabSwitching = function () {

  var tabsWrap = this.tabsBlock;
  tabsWrap.addEventListener("click", (e) =>{
    if (e.target.nodeName == "LABEL") {
      var clickedTabIndex = e.target.previousElementSibling.value;
      this.setCurrentSheet(clickedTabIndex);
      this.creatingTableDom(clickedTabIndex);
      this.parentBlock.scrollTop = this.scrollPosition();
      this.makeGrid();
    }

    if (e.target.className == "del") {
      var tabToDelete = [e.target.parentElement, e.target.parentElement.previousElementSibling];
      var clickedButtonIndex = parseInt(e.target.parentElement.previousElementSibling.value);
      this.tabDelete(clickedButtonIndex, tabToDelete);
    }

  });

};


MyExcel.prototype.fillCells = function (curSheet) {

  var self = this;

  //Навешиваем событие на пересчет
  self.formulaCells = {};
  self.linkCells = {};

  self.table.addEventListener("changePartOfFormula", function(){
    for (x in self.formulaCells) {
      var row = x,
          cell = parseInt(self.formulaCells[x]),
          text = self.sheetObject[self.currentSheet][row][cell];
      self.table.rows[row].cells[cell].textContent = self.formulaParse( text );
    }
  });

  if (!curSheet) {
    var currentSheet = self.tabsBlock.querySelector(":checked").value;
  } else { var currentSheet = curSheet }

  this.currentSheet = currentSheet;
  var cells = this.table.querySelectorAll("td");
  cells = Array.prototype.slice.call(cells);

  for (var cell in cells) {
    if (!cells.hasOwnProperty(cell)) continue;
    cells[cell].textContent = null;
  }

  for (var row in this.sheetObject[currentSheet]) {
    if (row != "settings") {
      var currentRow = this.sheetObject[currentSheet][row];
      for (var cell in this.sheetObject[currentSheet][row]) {
        if (!this.sheetObject[currentSheet][row].hasOwnProperty(cell)) continue;
        if (!this.table.rows[row] || !this.table.rows[row].cells[cell]) continue;//Не заполняем несуществующие ячейки
        var text = currentRow[cell];
        var cellText = this.table.rows[row].cells[cell].textContent;
        if (text.search(/^=/) != "-1") {//проверка не формула ли это
          this.table.rows[row].cells[cell].textContent = self.formulaParse( text );

          //добавляем ячейки со ссылками в специальный объект
          var formulaArray = text.match(/[a-z]{1,}\d{1,}/gi);
          if (formulaArray != null)
          formulaArray.forEach(function(i){
          var char = i.match(/[a-z]{1,}/i)[0],
              digit = i.match(/\d{1,}/i)[0];
            self.linkCells[parseInt( digit )] = self.charToIndex( char );
          });

          //добавляем ячейки с формулами в специальный объект
          self.formulaCells[row] = cell;

        } else {
          this.table.rows[row].cells[cell].textContent = text;
        };
      }
    }
  }

};

var widgetEvent = new CustomEvent("changePartOfFormula", {
  bubbles: true
});


MyExcel.prototype.setCurrentSheet = function (index) {

  for (var sheet in this.sheetObject) {
    if (!this.sheetObject.hasOwnProperty(sheet)) continue;
    this.sheetObject[sheet].settings.current = false;
  }

  this.sheetObject[index].settings.current = true;

  localStorage[this.tableName] = JSON.stringify(this.sheetObject);

};


MyExcel.prototype.deleteSheet = function (index) {

  delete this.sheetObject[index];
  this.updateSheet();

};


MyExcel.prototype.sheetCreate = function (index) {

  this.sheetObject[index] = 
  {
    "settings" : 
    {
      "name" : "sheet" + index,
      "current" : true
    }
  };
  this.setCurrentSheet(index);
  localStorage[this.tableName] = JSON.stringify(this.sheetObject);

};


MyExcel.prototype.tabGenerating = function () {

  var tabsWrap = this.tabsBlock;
  var button = document.createElement("BUTTON");
  button.textContent = "+";
  button.addEventListener("click", (e) => {
    this.tabCreate(undefined, undefined, undefined, true);
  });
  tabsWrap.appendChild(button);

  if (!this.sheetObject[0]) {//если sheetObject пуст, следовательно это первый запуск, создаем один таб
    this.tabCreate(0, "sheet0", true)
    this.sheetCreate(0);

  } else { //если в this.sheetObject есть данные о листах, строим табы в соответствии с this.sheetObject

    for (var sheet in this.sheetObject) {
      
      if (!this.sheetObject.hasOwnProperty(sheet)) {
        continue;
      }

      var sheetIndex = parseInt(sheet);
      var name = this.sheetObject[sheet].settings.name;
      var current = this.sheetObject[sheet].settings.current;
      this.tabCreate(sheetIndex, name, current);
    }

  };

  this.tabSwitching();

};


MyExcel.prototype.tabCreate = function (counter, name, ifCurrent, isNew) {

  var lastInput = this.tabsBlock.querySelector("input:last-of-type");
  if (counter == undefined) {
    if (lastInput != undefined) {
      var lastInputIndex = parseInt(lastInput.value);
      counter = ++lastInputIndex;
    };
  }
  if (name == undefined) {
    name = "sheet" + counter;
  }

  var tabsWrap = this.tabsBlock;

  var input = document.createElement("INPUT");
  input.id = this.tableName + "tab" + counter;
  input.type = "radio";
  input.name = this.tableName + "tab";
  input.value = counter;
  if (ifCurrent == true) {//установка активного таба
    input.checked = true;
  };

  var label = document.createElement("LABEL");
  label.setAttribute("for", this.tableName + "tab" + counter);
  label.textContent = name;

  var deleteTab = document.createElement("DIV");
  deleteTab.className = "del";
  deleteTab.textContent = "X";

  tabsWrap.appendChild(input);
  tabsWrap.appendChild(label);
  label.appendChild(deleteTab);

  if (isNew == true) {
    this.sheetCreate(counter);
    tabsWrap.querySelector("[type='radio'][value='" + counter + "']").checked = true;
    this.creatingTableDom();
  }

};


MyExcel.prototype.updateSheet = function () {

  localStorage[this.tableName] = JSON.stringify(this.sheetObject);

  // 1. Создаём новый объект XMLHttpRequest
  var xhr = new XMLHttpRequest();
  var body = "fileName=" + "files/" + this.tableName + ".json" + "&" + "object=" + JSON.stringify(this.sheetObject);
  // 2. Конфигурируем его: GET-запрос на URL 'phones.json'
  xhr.open('POST', 'create.php', false);
  xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  // 3. Отсылаем запрос
  try {
  xhr.send(body);
  } catch(err) {console.log("Данные на сервере не сохранены")};
  // 4. Если код ответа сервера не 200, то это ошибка
  if (xhr.status != 200) {

    // обработать ошибку

    } else {

    // вывести результат

  }

};


MyExcel.prototype.cellParse = function(indexOfCell) {

  var cellIndex = this.charToIndex( indexOfCell.match(/[a-z]{1,}/i)[0] );
  var rowIndex = indexOfCell.match(/\d{1,}/)[0];

  var string = this.sheetObject[this.currentSheet][rowIndex - 1][cellIndex];

  if (string == undefined) {return 0}

  if (string.search(/^=/) != "-1") {

    return this.formulaParse(string);

  } else {return string}

};


MyExcel.prototype.formulaParse = function (string) {

  var self = this;
  var result = 0;
  var lastOperand = "";
  var mathArray = string.match(/\d{1,}|\+|\-|\*|\/|\^|\w{1,}\d{1,}/g);
  var arrayForEval = [];

  mathArray.forEach(function(item){//преобразование ссылок в числа
    if (item.match(/[a-z]{1,}\d{1,}/i) != null) { var parsedItem = self.cellParse(item); arrayForEval.push( parsedItem ) } else {
      arrayForEval.push(item);
    }
  });

  result = eval(arrayForEval.toString().replace(/,/g," "));

  return result;

};


MyExcel.prototype.ifFormula = function (string) {

  if (string.search(/^=/) != "-1") {return true} else {return false}

}


MyExcel.prototype.scrollWidth = function (string) {

  var div = document.createElement('div');

  div.style.overflowY = 'scroll';
  div.style.width = '50px';
  div.style.height = '50px';

  div.style.visibility = 'hidden';

  document.body.appendChild(div);
  var scrollWidth = div.offsetWidth - div.clientWidth;
  document.body.removeChild(div);
  return scrollWidth;

}


MyExcel.prototype.scrollPosition = function (string) {

  return 0;

}