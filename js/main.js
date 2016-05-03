 "use strict";

function MyExcel() {

  //private

  var self = this;
  var sTable = {};
  var tableName , parentBlock, table, tabsBlock;

  var charArray = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  
  var indexToChar = (index) => {

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

  var charToIndex = (char) => {
    var index = 0, rank = 0, resultArray = [];
    var strArray = char.split("");
    strArray.forEach( (currentChar) => {
      charArray.forEach( (item, i) => {
        if (item == currentChar) {

          resultArray.push(i);
        }
      });
    } );
    console.log(resultArray);
    resultArray.forEach( (item, i, arr) => {
      var add = item;
      for (var j = 0; j < arr.length - i - 1; j++) {
        if (j == 0) {add++}
        add = 26 * add
        console.log(add)
      }
      console.dir(add)
      index = index + add;
    });
    console.dir(index)

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


    var viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
    var viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

    var heightOfTable = viewportHeight - parseInt(window.getComputedStyle(parentBlock).getPropertyValue('margin-top'));
    if (width == undefined)  {parentBlock.style.width  = (viewportWidth + "px");}
    if (height == undefined) {parentBlock.style.height = (heightOfTable + "px");}
    if (cellWidth == undefined) {cellWidth = 80;}
    if (cellHeight == undefined) {cellHeight = 20;}

    var rowsInitialAmount = Math.ceil(viewportHeight / cellHeight);
    var columnsInitialAmount = Math.ceil(viewportWidth / cellWidth);

    //создание DOM таблицы
    var creatingTableDom = (() => {

      var tableFragment = document.createDocumentFragment();
      table = document.createElement("TABLE");
      table.className = "super-table";
      tableFragment.appendChild(table);

      for (var i = 0; i < rowsInitialAmount; i++) {
        var row = table.insertRow();
        for (var j = 0; j < columnsInitialAmount; j++) {
          row.insertCell();
        }
      }
      parentBlock.appendChild(table);
      sTable.fillCells();//Заполняем ячейки
      this.cellSelect();
    })();

  };

  this.cellSelect = () => {
    table.addEventListener("click", (e) => {
      if (e.target.tagName == "TD") {

        var rowIndex = e.target.parentElement.rowIndex;
        var cellIndex = e.target.cellIndex;
        var targetCell = table.tBodies[0].rows[rowIndex].cells[cellIndex];
        var input = document.createElement("INPUT");

        self.selectedCell = indexToChar(cellIndex - 1) + (rowIndex - 1);
        if (rowIndex != 0 && cellIndex != 0) {//если клик не по первым рядам...

          targetCell.className += "hover";

          var inputCreating = (e) => {
            console.log("inputCreating")
            input.value = targetCell.textContent;
            targetCell.textContent = null;
            targetCell.appendChild(input);
            input.focus();
            window.removeEventListener("keydown", inputCreating);
          }

          window.addEventListener("keydown", inputCreating);

          input.onblur = () => {

            var currentSheet = tabsBlock.querySelector(":checked").value;

            if (!sTable.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
              sTable.sheetObject[currentSheet][rowIndex] = {};
            }

            sTable.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;

            sTable.updateSheet();

            input.parentElement.textContent = input.value;
            input.remove();
            input = null;
            targetCell.className = targetCell.className.replace(/\bhover\b/,'');

          };
        }
      }
    });
  }

  this.getCoords = () => {

  }

  this.makeGrid = (table) => {

    var tableRows = Array.prototype.slice.call(table.rows);
    var firstRowCells = Array.prototype.slice.call(tableRows[0].cells);

    for (var cell in firstRowCells.slice(1)) {
      var cellOffset = parseInt(cell) + 1;
      firstRowCells[cellOffset].textContent = charArray[cell];
    }

    for (var row in tableRows.slice(1)) {
      var rowOffset = parseInt(row) + 1;
      tableRows[rowOffset].firstChild.textContent = row;
    }
    //console.dir(table.rows);
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
      sTable.fillCells();
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
      sTable.fillCells();

    } else {

      tabsBlock.querySelector("#tab" + sheetToRedirect).checked = true;
      sTable.fillCells();
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
        sTable.fillCells(clickedTabIndex);

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

    var cells = table.querySelectorAll("td");
    cells = Array.prototype.slice.call(cells);
    for (var cell in cells) {
      cells[cell].textContent = null;
    }

    for (var row in sTable.sheetObject[currentSheet]) {
      if (row != "settings") {
        for (var cell in sTable.sheetObject[currentSheet][row]) {
          table.rows[row].cells[cell].textContent = sTable.sheetObject[currentSheet][row][cell];
        }
      }
    }
    self.makeGrid(table);
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
