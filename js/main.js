if (localStorage.getItem('sheets')) {
  var sheetObject = JSON.parse(localStorage.sheets);
} else {
  var sheetObject = {};
};


(function tabGenerating(){

  var tabsWrap = document.querySelector("#tabs-wrap");
  var button = document.createElement("BUTTON");
  button.textContent = "+";
  button.addEventListener("click", function(e){
    tabCreate(undefined, undefined, undefined, true);
  });
  tabsWrap.appendChild(button);

  if (!sheetObject[0]) {//если sheetObject пуст, следовательно это первый запуск, создаем один таб
    console.log("ОБЪЕКТА НЕТ!!!!!!");
    tabCreate(0, "sheet0", true)
    sheetCreate(0);

  } else { //если в sheetObject есть данные о листах, строим табы в соответствии с sheetObject

    for (var sheet in sheetObject) {
      var sheetIndex = parseInt(sheet)
      var name = sheetObject[sheet].settings.name;
      var current = sheetObject[sheet].settings.current;
      tabCreate(sheetIndex, name, current);
    }

  };

  tabSwitching();

})();


function tabCreate(counter, name, ifCurrent, isNew) {

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

  tabsWrap.appendChild(input);
  tabsWrap.appendChild(label);
  label.appendChild(deleteTab);

  if (isNew == true) {
    sheetCreate(counter);
    document.querySelector("[name='tab'][value='" + counter + "']").checked = true;
    fillCells();
  }

};


function tabDelete(index, elements) {

  //удаляем таб
  var tabsWrap = document.querySelector("#tabs-wrap");
  elements.forEach(function(item, i, arr){
    tabsWrap.removeChild(item);
  });

  //удаляем объект листа и обновляем localStorage
  deleteSheet(index);

  //устанавливаем нужный таб и перерисовываем страницу
  var sheetToRedirect = setReaddressSheet(index)

  if (sheetToRedirect == undefined) {// если undefined — значит это был единственный лист, и нужно создать новый

    tabCreate(0, "sheet0", true)
    sheetCreate(0);
    fillCells();

  } else {

    document.querySelector("#tabs-wrap #tab" + sheetToRedirect).checked = true;
    fillCells();

  }


};


function setReaddressSheet(index) {

  //выясняем на какой таб будет переадресация после удаления
  var previousElement;//кодовое число, будет использоваться в случае если это первый лист
  for (var sheet in sheetObject) {
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


function tabSwitching() {

  var tabsWrap = document.querySelector("#tabs-wrap");
  tabsWrap.addEventListener("click", function(e){
    if (e.target.nodeName == "LABEL") {

      var clickedTabIndex = e.target.previousElementSibling.value;
      setCurrentSheet(clickedTabIndex);
      fillCells(clickedTabIndex);

    }
    console.dir(e.target);
    if (e.target.className == "del") {

      var tabToDelete = [e.target.parentElement, e.target.parentElement.previousElementSibling]
      var clickedButtonIndex = parseInt(e.target.parentElement.previousElementSibling.value);
      tabDelete(clickedButtonIndex, tabToDelete);

    }
  });
};

//заполнение ячеек сохраненными данными
function fillCells(currentSheet) {

  if (currentSheet == undefined) {
    var currentSheet = document.querySelector("[name='tab']:checked").value;
  }

  var cells = document.querySelectorAll("td");
  cells = Array.prototype.slice.call(cells);
  for (var cell in cells) {
    cells[cell].textContent = null;
  }
  var table = document.getElementById("super-table");

  for (var row in sheetObject[currentSheet]) {
    if (row != "settings") {
      for (var cell in sheetObject[currentSheet][row]) {
        table.rows[row].cells[cell].textContent = sheetObject[currentSheet][row][cell];
      }
    }
  }
};


tableInit("table-wrap");
function tableInit(block, width, height, cellWidth, cellHeight){

  var charArray = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];
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
    tableFragment.appendChild(table)

    for (var i = 0; i < rowsInitialAmount; i++) {
      var row = table.insertRow();
      for (var j = 0; j < columnsInitialAmount; j++) {
        row.insertCell();
      }
    }
    parentBlock.appendChild(table);
    fillCells();//Заполняем ячейки
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

          if (!sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
            sheetObject[currentSheet][rowIndex] = {};
          }

          sheetObject[currentSheet][rowIndex][cellIndex] = input.value;

          updateSheet();

          input.parentElement.textContent = input.value;
          input.remove();
          input = null;

        };
      }
    });

  })();

};


function setCurrentSheet(index) {

  for (var sheet in sheetObject) {
    sheetObject[sheet].settings.current = false
  }

  sheetObject[index].settings.current = true;

  localStorage.sheets = JSON.stringify(sheetObject);

};


function deleteSheet(index) {

  delete sheetObject[index];
  updateSheet();

};


function sheetCreate(index) {

  sheetObject[index] = 
  {
    "settings" : 
    {
      "name" : "sheet" + index,
      "current" : true
    }
  };
  setCurrentSheet(index);
  localStorage.sheets = JSON.stringify(sheetObject);

};


function updateSheet() {

  localStorage.sheets = JSON.stringify(sheetObject);

};