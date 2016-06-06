 "use strict";

function MyExcel() {

  var self = this;

  this.viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  this.viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

  this.charArray = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

};


MyExcel.prototype.init = function (block, width, height, cellWidth, cellHeight) {

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

  var heightOfTable = this.viewportHeight - parseInt(window.getComputedStyle(this.parentBlock).getPropertyValue('margin-top'));
  if (width == undefined)  {this.parentBlock.style.width  = (this.viewportWidth + "px");}
  if (height == undefined) {this.parentBlock.style.height = (heightOfTable + "px");}
  if (cellWidth == undefined) {self.cellWidth = 80;}
  if (cellHeight == undefined) {self.cellHeight = 20;}

  self.creatingTableDom();

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
          row.insertCell();
        }
      }

    }

    if (self.parentBlock.scrollLeft == self.parentBlock.scrollWidth - self.parentBlock.clientWidth) {

      var numberOfAddingColumns = 5;
      //добавления в таблицу разметки
      for (var i = 0; i < numberOfAddingColumns; i++) {
        var cell = self.gridTableElementColumns.rows[0].insertCell();
        var cellIndex = parseInt(cell.cellIndex);
        self.gridTableElementColumns.rows[0].cells[cellIndex].textContent = self.indexToChar(cellIndex);
      }
      //добавления в основную таблицу
      for (var x in self.table.rows) {
        if (!self.table.rows.hasOwnProperty(x)) continue;
        for (var i = 0; i < numberOfAddingColumns; i++) {
          var cell = self.table.rows[x].insertCell();
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


MyExcel.prototype.creatingTableDom = function (currentSheet) {

  //создание DOM таблицы
  if (currentSheet == undefined) {
    var currentSheet = this.tabsBlock.querySelector(":checked").value;
  }

  var maxRowNumber  = 0;
  var maxCellNumber = 0;

  for (var row in this.sheetObject[currentSheet]) {

    if (!this.sheetObject[currentSheet].hasOwnProperty(row)) continue;
    var currentRowNumber = parseInt( row );
    if ( currentRowNumber > maxRowNumber ) { maxRowNumber = currentRowNumber };

    for ( var cell in this.sheetObject[currentSheet][row]) {

      if (!this.sheetObject[currentSheet][row].hasOwnProperty(cell)) continue;
      var currentCellNumber = parseInt( cell );
      if ( currentCellNumber > maxCellNumber ) { maxCellNumber = currentCellNumber };

    }

  }

  if (this.table) {this.parentBlock.removeChild(this.table);}

  var rowsInitialAmount = Math.ceil(this.viewportHeight / this.cellHeight);
  var columnsInitialAmount = Math.ceil(this.viewportWidth / this.cellWidth);

  if (rowsInitialAmount < maxRowNumber) {rowsInitialAmount = maxRowNumber + 1};
  if (columnsInitialAmount < maxCellNumber) {columnsInitialAmount = maxCellNumber + 1};

  this.globalInput = document.createElement("INPUT");
  this.globalInput.id = this.tableName + "-global-input";
  this.globalInput.className = "global-input";
  this.parentBlock.appendChild(this.globalInput);

  var tableFragment = document.createDocumentFragment();
  this.table = document.createElement("TABLE");
  this.table.className = "super-table";
  tableFragment.appendChild(this.table);

  var tt = '';
  var _ = Date.now();
  for (var i = 0; i < rowsInitialAmount; i++) {
  	tt += '<tr>';
    for (var j = 0; j < columnsInitialAmount; j++) {
      tt += '<td><td>';
    }
    tt += '</tr>';
  }  
  this.table.innerHTML = tt;
  alert(Date.now() - _)
/*
  for (var i = 0; i < rowsInitialAmount; i++) {
    var row = this.table.insertRow();
    for (var j = 0; j < columnsInitialAmount; j++) {
      row.insertCell();
    }
  }
*/
  this.parentBlock.appendChild(this.table);

  this.fillCells(currentSheet);//Заполняем ячейки
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
          input.value = input.value + self.selectedCellCoords;
          input.focus();
        }

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
          deleteHoverClass();
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

        var currentSheet = self.tabsBlock.querySelector(":checked").value;

        if (!self.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
          self.sheetObject[currentSheet][rowIndex] = {};
        }

        self.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;
        self.updateSheet();

        if (self.ifFormula(input.value)) {
          input.parentElement.textContent = self.formulaParse( input.value )
        } else {
          input.parentElement.textContent = input.value
        }

        input = null;

        self.table.removeEventListener("click", addReference);
        self.table.removeEventListener("click", deleteHoverClass);
        self.table.addEventListener("click", self.hoverFunction);
        inputUnfocus();
      };


      function inputUnfocus () {

        self.globalInput.className = self.globalInput.className.replace(/\bactive\b/,'');

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

    }
    if (e.target.className == "del") {

      var tabToDelete = [e.target.parentElement, e.target.parentElement.previousElementSibling];
      var clickedButtonIndex = parseInt(e.target.parentElement.previousElementSibling.value);
      this.tabDelete(clickedButtonIndex, tabToDelete);

    }
  });

};


MyExcel.prototype.fillCells = function (currentSheet) {

  var self = this;

  if (currentSheet == undefined) {
    var currentSheet = tabsBlock.querySelector(":checked").value;
  }
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
        var text = currentRow[cell];
        var cellText = this.table.rows[row].cells[cell].textContent;
        if (text.search(/^=/) != "-1") {//проверка не формула ли это
          this.table.rows[row].cells[cell].textContent = self.formulaParse( text )
        } else {
          this.table.rows[row].cells[cell].textContent = text;
        };
      }
    }
  }

};


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
  console.log()
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