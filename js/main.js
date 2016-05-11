 "use strict";

function MyExcel() {

  //private

  var self = this;
  var sTable = {};
  var tableName, parentBlock, tabsBlock, gridTableElementRows, gridTableElementColumns;

  var viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  var viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

  var charArray = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  
  this.indexToChar = (index) => {

    var word = "", result;

    var oneMoreTime = (lastResult) => {

      if (lastResult == undefined) {lastResult = index}
      result = Math.floor(lastResult / 26);
      if (result > 26) {
        var middleIndex = result % 26;
        word = word + charArray[middleIndex - 1];
        oneMoreTime(result)
      } else {
        word = charArray[result - 1] + word;
        if (result > 26) { oneMoreTime(result) } else { theLastChar(index % 26) };
      }

    }

    var theLastChar = (result) => {

      if (result == undefined) {result = index} 
      word = word + charArray[result];
    }

    if (index < 26) {theLastChar()} else {oneMoreTime()}
 
    return word;
  };

  this.charToIndex = function (char) {

    var index = 0, rank = 0, resultArray = [];
    var strArray = char.split("");

    strArray.forEach( (currentChar) => {
      var indexOfCurrentChar = charArray.indexOf(currentChar.toUpperCase());
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

  //public

  this.init = (block, width, height, cellWidth, cellHeight) => {

    tableName = block;//определяем имя таблицы. Объект в Local Storage будет называться так же
    parentBlock = document.getElementById(block);

    if (localStorage.getItem(tableName)) {
      sTable.sheetObject = JSON.parse(localStorage[tableName]);
    } else {
      sTable.sheetObject = {};
    };

    tabsBlock = document.createElement("DIV");
    tabsBlock.className = "tabs-wrap";
    parentBlock.parentNode.insertBefore(tabsBlock, parentBlock.nextSibling);
    sTable.tabGenerating();

    var heightOfTable = viewportHeight - parseInt(window.getComputedStyle(parentBlock).getPropertyValue('margin-top'));
    if (width == undefined)  {parentBlock.style.width  = (viewportWidth + "px");}
    if (height == undefined) {parentBlock.style.height = (heightOfTable + "px");}
    if (cellWidth == undefined) {self.cellWidth = 80;}
    if (cellHeight == undefined) {self.cellHeight = 20;}

    self.creatingTableDom();
    (function (){
      parentBlock.addEventListener("scroll", function(){

        if (parentBlock.scrollTop == parentBlock.scrollHeight - parentBlock.clientHeight) {

          var cellsQuantity = self.table.rows[0].children.length;

          for (var i = 0; i < 5; i++) {
            var row = self.table.insertRow();
            //добавления в таблицу разметки
            var rowGrid = gridTableElementRows.insertRow();
            rowGrid.insertCell();
            gridTableElementRows.rows[row.rowIndex].cells[0].textContent = row.rowIndex + 1;

            for (var j = 0; j < cellsQuantity; j++) {
              row.insertCell();
            }
          }

        }

        if (parentBlock.scrollLeft == parentBlock.scrollWidth - parentBlock.clientWidth) {

          var numberOfAddingColumns = 5;
          //добавления в таблицу разметки
          for (var i = 0; i < numberOfAddingColumns; i++) {
            var cell = gridTableElementColumns.rows[0].insertCell();
            var cellIndex = parseInt(cell.cellIndex);
            gridTableElementColumns.rows[0].cells[cellIndex].textContent = self.indexToChar(cellIndex);
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
    })();
  };

    //создание DOM таблицы
  this.creatingTableDom = (currentSheet) => {

    if (currentSheet == undefined) {
      var currentSheet = tabsBlock.querySelector(":checked").value;
    }

    var maxRowNumber  = 0;
    var maxCellNumber = 0;

    for (var row in sTable.sheetObject[currentSheet]) {

      var currentRowNumber = parseInt( row );
      if ( currentRowNumber > maxRowNumber ) { maxRowNumber = currentRowNumber };

      for ( var cell in sTable.sheetObject[currentSheet][row]) {

        var currentCellNumber = parseInt( cell );
        if ( currentCellNumber > maxCellNumber ) { maxCellNumber = currentCellNumber };

      }

    }

    if (self.table) {parentBlock.removeChild(self.table);}

    var rowsInitialAmount = Math.ceil(viewportHeight / self.cellHeight);
    var columnsInitialAmount = Math.ceil(viewportWidth / self.cellWidth);

    if (rowsInitialAmount < maxRowNumber) {rowsInitialAmount = maxRowNumber + 1};
    if (columnsInitialAmount < maxCellNumber) {columnsInitialAmount = maxCellNumber + 1};

    var tableFragment = document.createDocumentFragment();
    self.table = document.createElement("TABLE");
    self.table.className = "super-table";
    tableFragment.appendChild(self.table);

    for (var i = 0; i < rowsInitialAmount; i++) {
      var row = self.table.insertRow();
      for (var j = 0; j < columnsInitialAmount; j++) {
        row.insertCell();
      }
    }
    parentBlock.appendChild(self.table);

    sTable.fillCells(currentSheet);//Заполняем ячейки
    this.cellSelect();
  };

  this.cellSelect = () => {
    self.table.addEventListener("click", (e) => {
      if (e.target.tagName == "TD") {

        var rowIndex = e.target.parentElement.rowIndex;
        var cellIndex = e.target.cellIndex;
        var targetCell = self.table.tBodies[0].rows[rowIndex].cells[cellIndex];
        var input = document.createElement("INPUT");

        self.selectedCell = self.indexToChar(cellIndex - 1) + (rowIndex - 1);

        targetCell.className += "hover";

        var inputCreating = (e) => {
          input.value = targetCell.textContent;
          targetCell.textContent = null;
          targetCell.appendChild(input);
          input.focus();
          window.removeEventListener("keydown", inputCreating);
        }

        window.addEventListener("keydown", inputCreating);

        self.table.addEventListener("click", deleteHoverClass);

        function deleteHoverClass(){
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

          var currentSheet = tabsBlock.querySelector(":checked").value;

          if (!sTable.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
            sTable.sheetObject[currentSheet][rowIndex] = {};
          }

          sTable.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;
          sTable.updateSheet();
          input.parentElement.textContent = input.value;
          input = null;

        };

      }
    });
  }

  this.getCoords = () => {

  }

  this.makeGrid = () => {

    var self = this;

    this.addNewRow = (numberOfRow) => {

      gridTableElementRows.rows[numberOfRow].cells[0].textContent = numberOfRow + 1;

    }

    this.addNewColumn = (numberOfColumn) => {

      gridTableElementColumns.rows[0].cells[numberOfColumn].textContent = self.indexToChar(numberOfColumn);

    }

    var tableRows = Array.prototype.slice.call(self.table.rows);
    var firstRowCells = Array.prototype.slice.call(tableRows[0].cells);

    //таблица для строк
    gridTableElementRows = document.createElement("TABLE");
    gridTableElementRows.className = "super-table-rows";
    parentBlock.appendChild(gridTableElementRows);

    for (var i = 0; i < tableRows.length; i++) {
      var row = gridTableElementRows.insertRow();
      row.insertCell();
      this.addNewRow(i);
    }

    //таблица для столбцов
    gridTableElementColumns = document.createElement("TABLE");
    gridTableElementColumns.className = "super-table-columns";
    parentBlock.appendChild(gridTableElementColumns);

    var row = gridTableElementColumns.insertRow();

    for (var i = 0; i < firstRowCells.length; i++) {
      row.insertCell();
      this.addNewColumn(i);
    }

    (function (){
      parentBlock.addEventListener("scroll", function(){

        gridTableElementRows.style.left = parentBlock.scrollLeft + "px";

        gridTableElementColumns.style.top = parentBlock.scrollTop + "px";

      });
    })();

  };


  sTable.tabGenerating = () => {

    var tabsWrap = tabsBlock;
    var button = document.createElement("BUTTON");
    button.textContent = "+";
    button.addEventListener("click", (e) => {
      sTable.tabCreate(undefined, undefined, undefined, true);
    });
    tabsWrap.appendChild(button);

    if (!sTable.sheetObject[0]) {//если sTable.sheetObject пуст, следовательно это первый запуск, создаем один таб
      sTable.tabCreate(0, "sheet0", true)
      sTable.sheetCreate(0);

    } else { //если в sTable.sheetObject есть данные о листах, строим табы в соответствии с sTable.sheetObject

      for (var sheet in sTable.sheetObject) {
        if (!sTable.sheetObject.hasOwnProperty(sheet)) {
          continue;
        }

        var sheetIndex = parseInt(sheet);
        var name = sTable.sheetObject[sheet].settings.name;
        var current = sTable.sheetObject[sheet].settings.current;
        sTable.tabCreate(sheetIndex, name, current);
      }

    };

    sTable.tabSwitching();

  };


  sTable.tabCreate = (counter, name, ifCurrent, isNew) => {

    var lastInput = tabsBlock.querySelector("input:last-of-type");
    if (counter == undefined) {
      if (lastInput != undefined) {
        var lastInputIndex = parseInt(lastInput.value);
        counter = ++lastInputIndex;
      };
    }
    if (name == undefined) {
      name = "sheet" + counter;
    }

    var tabsWrap = tabsBlock;

    var input = document.createElement("INPUT");
    input.id = tableName + "tab" + counter;
    input.type = "radio";
    input.name = tableName + "tab";
    input.value = counter;
    if (ifCurrent == true) {//установка активного таба
      input.checked = true;
    };

    var label = document.createElement("LABEL");
    label.setAttribute("for", tableName + "tab" + counter);
    label.textContent = name;

    var deleteTab = document.createElement("DIV");
    deleteTab.className = "del";
    deleteTab.textContent = "X";

    tabsWrap.appendChild(input);
    tabsWrap.appendChild(label);
    label.appendChild(deleteTab);

    if (isNew == true) {
      sTable.sheetCreate(counter);
      tabsBlock.querySelector("[type='radio'][value='" + counter + "']").checked = true;
      self.creatingTableDom();
    }

  };


  sTable.tabDelete = (index, elements) => {

    //удаляем таб
    var tabsWrap = tabsBlock;
    elements.forEach((item, i, arr) => {
      tabsWrap.removeChild(item);
    });

    //удаляем объект листа и обновляем localStorage
    sTable.deleteSheet(index);

    //устанавливаем нужный таб и перерисовываем страницу
    var sheetToRedirect = sTable.setReaddressSheet(index);

    if (sheetToRedirect == undefined) {// если undefined — значит это был единственный лист, и нужно создать новый

      sTable.tabCreate(0, "sheet0", true);
      sTable.sheetCreate(0);
      self.creatingTableDom();

    } else {

      tabsBlock.querySelector("#" + tableName + "tab" + sheetToRedirect).checked = true;
      self.creatingTableDom();
      sTable.setCurrentSheet(sheetToRedirect);

    }

  };


  sTable.setReaddressSheet = (index) => {

    //выясняем на какой таб будет переадресация после удаления
    var previousElement;//кодовое число, будет использоваться в случае если это первый лист
    for (var sheet in sTable.sheetObject) {
        if (previousElement != undefined) {//если первый лист, проверяем не единственный ли это лист
          if (sheet != index) {
            return sheet;
            break;
          }
          return previousElement;
          break;
        }
      previousElement = sheet;
    }

  };


  sTable.tabSwitching = () => {

    var tabsWrap = tabsBlock;
    tabsWrap.addEventListener("click", (e) =>{
      if (e.target.nodeName == "LABEL") {

        var clickedTabIndex = e.target.previousElementSibling.value;
        sTable.setCurrentSheet(clickedTabIndex);
        self.creatingTableDom(clickedTabIndex);

      }
      if (e.target.className == "del") {

        var tabToDelete = [e.target.parentElement, e.target.parentElement.previousElementSibling];
        var clickedButtonIndex = parseInt(e.target.parentElement.previousElementSibling.value);
        sTable.tabDelete(clickedButtonIndex, tabToDelete);

      }
    });

  };


  sTable.fillCells = (currentSheet) => {
    if (currentSheet == undefined) {
      var currentSheet = tabsBlock.querySelector(":checked").value;
    }

    var cells = self.table.querySelectorAll("td");
    cells = Array.prototype.slice.call(cells);
    for (var cell in cells) {
      cells[cell].textContent = null;
    }

    for (var row in sTable.sheetObject[currentSheet]) {
      if (row != "settings") {
        for (var cell in sTable.sheetObject[currentSheet][row]) {
          self.table.rows[row].cells[cell].textContent = sTable.sheetObject[currentSheet][row][cell];
        }
      }
    }
    self.makeGrid(self.table);
  };


  sTable.setCurrentSheet = (index) => {

    for (var sheet in sTable.sheetObject) {
      sTable.sheetObject[sheet].settings.current = false;
    }

    sTable.sheetObject[index].settings.current = true;

    localStorage[tableName] = JSON.stringify(sTable.sheetObject);

  };


  sTable.deleteSheet = (index) => {

    delete sTable.sheetObject[index];
    sTable.updateSheet();

  };


  sTable.sheetCreate = (index) => {

    sTable.sheetObject[index] = 
    {
      "settings" : 
      {
        "name" : "sheet" + index,
        "current" : true
      }
    };
    sTable.setCurrentSheet(index);
    localStorage[tableName] = JSON.stringify(sTable.sheetObject);

  };


  sTable.updateSheet = () => {

    localStorage[tableName] = JSON.stringify(sTable.sheetObject);

  };

}
MyExcel.prototype.cellParse = function(indexOfCell) {

  var cellIndex = this.charToIndex( indexOfCell.match(/[a-z]{1,}/i)[0] );
  var rowIndex = indexOfCell.match(/\d{1,}/)[0];

  var needlyCell = this.table.rows[rowIndex - 1].cells[cellIndex];
  var string = needlyCell.textContent;

  if (string.search(/^=/) != "-1") {

    var that = this;
    var result = 0;


    var lastOperand = "";
    var mathArray = string.match(/\d{1,}|\+|\-|\*|\/|\^|\w{1,}\d{1,}/g);

    mathArray.forEach(function(item){

      if (item.match(/\+|\-|\*|\/|\^/) != null) {//если оператор
        lastOperand = item;
      }

      if (item.match(/^\d{1,}/) != null) {//если число
        doOperation(parseInt(item));
      }

      if (item.match(/[a-z]{1,}\d{1,}/i) != null) {//если ссылка
        var referenceCellContent = that.cellParse(item);
        doOperation(parseInt(referenceCellContent));
      }

    });
    return result;
  } else {return string}

  function doOperation(number) {
    if (lastOperand == "" || lastOperand == undefined) {lastOperand = "+"}//возникает только в случае первого запуска, или ошибочного ввода.
    result = eval(result + lastOperand + number);
  };
}