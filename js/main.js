if (localStorage.getItem('sheets')) {
  var sheetObject = JSON.parse(localStorage.sheets);
} else {
  var sheetObject = {};
};

//заполнение ячеек сохраненными данными
function fillCells(){
  var cells = document.querySelectorAll("td");
  cells = Array.prototype.slice.call(cells);
  for (var cell in cells) {
    cells[cell].textContent = null;
  }
  var table = document.getElementById("super-table");
  var currentSheet = document.querySelectorAll("[name='tab']:checked")[0].value;
  for (var row in sheetObject[currentSheet]) {
    for (var cell in sheetObject[currentSheet][row]) {
      table.rows[row].cells[cell].textContent = sheetObject[currentSheet][row][cell];
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

          var currentSheet = document.querySelectorAll("[name='tab']:checked")[0].value;

          if (!sheetObject.hasOwnProperty(currentSheet)) {//объект sheetObject представляет из себя трехмерный массив, проврки нужны чтобы проверить существует ли уже нужный объект
            sheetObject[currentSheet] = {};
          }
          if (!sheetObject[currentSheet].hasOwnProperty(rowIndex)) {
            sheetObject[currentSheet][rowIndex] = {};
          }
          sheetObject[currentSheet][rowIndex][cellIndex] = input.value;

          localStorage.sheets = JSON.stringify(sheetObject);

          input.parentElement.textContent = input.value;
          input.remove();
          input = null;

        };
      }
    });

  })();

};
(function tabSwitching(){
  var tabs = document.querySelectorAll("[name='tab']");
  tabs = Array.prototype.slice.call(tabs);
  for (var tab in tabs) {
    tabs[tab].addEventListener("click", function(e){
      fillCells();
    })
  }
})();