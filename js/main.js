 "use strict";

function MyExcel() {

  var self = this;

  this.viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  this.viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

  this.charArray = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

}


MyExcel.prototype.init = function (block, width, height, cellWidth, cellHeight) {

  var self = this;

  self.tableName = block;//определяем имя таблицы. Объект в Local Storage будет называться так же
  this.parentBlock = document.getElementById(block);

  if (localStorage.getItem(self.tableName)) {
    self.sheetObject = JSON.parse(localStorage[self.tableName]);
  } else {
    self.sheetObject = {};
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

    var currentRowNumber = parseInt( row );
    if ( currentRowNumber > maxRowNumber ) { maxRowNumber = currentRowNumber };

    for ( var cell in this.sheetObject[currentSheet][row]) {

      var currentCellNumber = parseInt( cell );
      if ( currentCellNumber > maxCellNumber ) { maxCellNumber = currentCellNumber };

    }

  }

  if (this.table) {this.parentBlock.removeChild(this.table);}

  var rowsInitialAmount = Math.ceil(this.viewportHeight / this.cellHeight);
  var columnsInitialAmount = Math.ceil(this.viewportWidth / this.cellWidth);

  if (rowsInitialAmount < maxRowNumber) {rowsInitialAmount = maxRowNumber + 1};
  if (columnsInitialAmount < maxCellNumber) {columnsInitialAmount = maxCellNumber + 1};

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
  this.fillCells(currentSheet);//Заполняем ячейки
  this.cellSelect();

};


MyExcel.prototype.cellSelect = function () {

  var self = this;

  this.table.addEventListener("click", (e) => {
    if (e.target.tagName == "TD") {

      var rowIndex = e.target.parentElement.rowIndex;
      var cellIndex = e.target.cellIndex;
      var targetCell = this.table.tBodies[0].rows[rowIndex].cells[cellIndex];
      var input = document.createElement("INPUT");

      this.selectedCell = this.indexToChar(cellIndex - 1) + (rowIndex - 1);

      targetCell.className += "hover";

      function inputCreating (e) {
        input.value = targetCell.textContent;
        targetCell.textContent = null;
        targetCell.appendChild(input);
        input.focus();
        window.removeEventListener("keydown", inputCreating);
      }

      window.addEventListener("keydown", inputCreating);

      self.table.addEventListener("click", deleteHoverClass);

      function deleteHoverClass() {
        targetCell.className = targetCell.className.replace(/\bhover\b/,'');
        self.table.removeEventListener("click", deleteHoverClass);
      }

      input.addEventListener("blur", saveValue)

      input.addEventListener("keypress", function(e){

        var key = e.which || e.keyCode;
        if (key === 13) { 
          saveValue();
        }

      });

      function saveValue() {
        console.log("!!!")
        var currentSheet = self.tabsBlock.querySelector(":checked").value;

        if (!self.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
          self.sheetObject[currentSheet][rowIndex] = {};
        }

        self.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;
        self.updateSheet();
        input.parentElement.textContent = input.value;
        input = null;

      };

    }
  });

}


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
      console.log(clickedButtonIndex);
      this.tabDelete(clickedButtonIndex, tabToDelete);

    }
  });

};


MyExcel.prototype.fillCells = function (currentSheet) {

  if (currentSheet == undefined) {
    var currentSheet = tabsBlock.querySelector(":checked").value;
  }

  var cells = this.table.querySelectorAll("td");
  cells = Array.prototype.slice.call(cells);

  for (var cell in cells) {
    cells[cell].textContent = null;
  }

  for (var row in this.sheetObject[currentSheet]) {
    if (row != "settings") {
      for (var cell in this.sheetObject[currentSheet][row]) {
        this.table.rows[row].cells[cell].textContent = this.sheetObject[currentSheet][row][cell];
      }
    }
  }

  this.makeGrid(this.table);

};


MyExcel.prototype.setCurrentSheet = function (index) {

  for (var sheet in this.sheetObject) {
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

};


MyExcel.prototype.cellParse = function(indexOfCell) {

  var cellIndex = this.charToIndex( indexOfCell.match(/[a-z]{1,}/i)[0] );
  var rowIndex = indexOfCell.match(/\d{1,}/)[0];

  var needlyCell = this.table.rows[rowIndex - 1].cells[cellIndex];
  var string = needlyCell.textContent;

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
    if (item.match(/[a-z]{1,}\d{1,}/i) != null) { arrayForEval.push( self.cellParse(item) ) } else {
      arrayForEval.push(item);
    }
  });

  result = eval(arrayForEval.toString().replace(/,/g," "));

  return result;

};