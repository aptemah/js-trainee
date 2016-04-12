var sTable = {};


  if (localStorage.getItem('sheets')) {
    sTable.sheetObject = JSON.parse(localStorage.sheets);
  } else {
    sTable.sheetObject = {};
  };


  sTable.charArray = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];


  sTable.init = function (block, width, height, cellWidth, cellHeight) {

    sTable.tabGenerating();

    var parentBlock = document.getElementById(block);

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
    var creatingTableDom = (function() {

      var tableFragment = document.createDocumentFragment();
      var table = document.createElement("TABLE");
      table.id = "super-table";
      tableFragment.appendChild(table);

      for (var i = 0; i < rowsInitialAmount; i++) {
        var row = table.insertRow();
        for (var j = 0; j < columnsInitialAmount; j++) {
          row.insertCell();
        }
      }
      parentBlock.appendChild(table);
      sTable.fillCells();//Заполняем ячейки
      //Навешиваем события на клик и блур
      table.addEventListener("click", function(e){
        if (e.target.tagName == "TD") {

          var rowIndex = e.target.parentElement.rowIndex;
          var cellIndex = e.target.cellIndex;
          var targetCell = table.tBodies[0].rows[rowIndex].cells[cellIndex];
          var input = document.createElement("INPUT");

          input.value = targetCell.textContent;
          targetCell.textContent = null;
          targetCell.appendChild(input);
          input.focus();

          input.onblur = function(){

            var currentSheet = document.querySelector("[name='tab']:checked").value;

            if (!sTable.sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
              sTable.sheetObject[currentSheet][rowIndex] = {};
            }

            sTable.sheetObject[currentSheet][rowIndex][cellIndex] = input.value;

            sTable.updateSheet();

            input.parentElement.textContent = input.value;
            input.remove();
            input = null;

          };
        }
      });

      sTable.makeGrid(table);

    })();

  };


  sTable.makeGrid = function(table) {

    var tableRows = Array.prototype.slice.call(table.rows);
    var firstRowCells = Array.prototype.slice.call(tableRows[0].cells);

    for (var cell in firstRowCells.slice(1)) {
      var cellOffset = parseInt(cell) + 1;
      firstRowCells[cellOffset].textContent = sTable.charArray[cell];
    }

    for (var row in tableRows.slice(1)) {
      var rowOffset = parseInt(row) + 1;
      tableRows[rowOffset].firstChild.textContent = row;
    }
    //console.dir(table.rows);
  };


  sTable.tabGenerating = function() {

    var tabsWrap = document.querySelector("#tabs-wrap");
    var button = document.createElement("BUTTON");
    button.textContent = "+";
    button.addEventListener("click", function(e){
      sTable.tabCreate(undefined, undefined, undefined, true);
    });
    tabsWrap.appendChild(button);

    if (!sTable.sheetObject[0]) {//если sTable.sheetObject пуст, следовательно это первый запуск, создаем один таб
      sTable.tabCreate(0, "sheet0", true)
      sTable.sheetCreate(0);

    } else { //если в sTable.sheetObject есть данные о листах, строим табы в соответствии с sTable.sheetObject

      for (var sheet in sTable.sheetObject) {
        var sheetIndex = parseInt(sheet);
        var name = sTable.sheetObject[sheet].settings.name;
        var current = sTable.sheetObject[sheet].settings.current;
        sTable.tabCreate(sheetIndex, name, current);
      }

    };

    sTable.tabSwitching();

  };


  sTable.tabCreate = function(counter, name, ifCurrent, isNew) {

    var lastInput = document.querySelector("#tabs-wrap input:last-of-type");
    if (counter == undefined) {
      if (lastInput != undefined) {
        var lastInputIndex = parseInt(lastInput.value);
        counter = ++lastInputIndex;
      };
    }
    if (name == undefined) {
      name = "sheet" + counter;
    }

    var tabsWrap = document.querySelector("#tabs-wrap");

    var input = document.createElement("INPUT");
    input.id = "tab" + counter;
    input.type = "radio";
    input.name = "tab";
    input.value = counter;
    if (ifCurrent == true) {//установка активного таба
      input.checked = true;
    };

    var label = document.createElement("LABEL");
    label.setAttribute("for", "tab" + counter);
    label.textContent = name;

    var deleteTab = document.createElement("DIV");
    deleteTab.className = "del";
    deleteTab.textContent = "X";

    tabsWrap.appendChild(input);
    tabsWrap.appendChild(label);
    label.appendChild(deleteTab);

    if (isNew == true) {
      sTable.sheetCreate(counter);
      document.querySelector("[name='tab'][value='" + counter + "']").checked = true;
      sTable.fillCells();
    }

  };


  sTable.tabDelete = function(index, elements) {

    //удаляем таб
    var tabsWrap = document.querySelector("#tabs-wrap");
    elements.forEach(function(item, i, arr){
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

      document.querySelector("#tabs-wrap #tab" + sheetToRedirect).checked = true;
      sTable.fillCells();
      sTable.setCurrentSheet(sheetToRedirect);

    }

  };


  sTable.setReaddressSheet = function(index) {

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


  sTable.tabSwitching = function() {

    var tabsWrap = document.querySelector("#tabs-wrap");
    tabsWrap.addEventListener("click", function(e){
      if (e.target.nodeName == "LABEL") {

        var clickedTabIndex = e.target.previousElementSibling.value;
        sTable.setCurrentSheet(clickedTabIndex);
        sTable.fillCells(clickedTabIndex);

      }
      console.dir(e.target);
      if (e.target.className == "del") {

        var tabToDelete = [e.target.parentElement, e.target.parentElement.previousElementSibling]
        var clickedButtonIndex = parseInt(e.target.parentElement.previousElementSibling.value);
        sTable.tabDelete(clickedButtonIndex, tabToDelete);

      }
    });

  };


  sTable.fillCells = function(currentSheet) {

    if (currentSheet == undefined) {
      var currentSheet = document.querySelector("[name='tab']:checked").value;
    }

    var cells = document.querySelectorAll("td");
    cells = Array.prototype.slice.call(cells);
    for (var cell in cells) {
      cells[cell].textContent = null;
    }
    var table = document.getElementById("super-table");

    for (var row in sTable.sheetObject[currentSheet]) {
      if (row != "settings") {
        for (var cell in sTable.sheetObject[currentSheet][row]) {
          table.rows[row].cells[cell].textContent = sTable.sheetObject[currentSheet][row][cell];
        }
      }
    }
  };


  sTable.setCurrentSheet = function(index) {

    for (var sheet in sTable.sheetObject) {
      sTable.sheetObject[sheet].settings.current = false;
    }

    sTable.sheetObject[index].settings.current = true;

    localStorage.sheets = JSON.stringify(sTable.sheetObject);

  };


  sTable.deleteSheet = function(index) {

    delete sTable.sheetObject[index];
    sTable.updateSheet();

  };


  sTable.sheetCreate = function(index) {

    sTable.sheetObject[index] = 
    {
      "settings" : 
      {
        "name" : "sheet" + index,
        "current" : true
      }
    };
    sTable.setCurrentSheet(index);
    localStorage.sheets = JSON.stringify(sTable.sheetObject);

  };


  sTable.updateSheet = function() {

    localStorage.sheets = JSON.stringify(sTable.sheetObject);

  };

document.addEventListener("DOMContentLoaded", sTable.init("table-wrap"));