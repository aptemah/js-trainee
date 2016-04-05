tableInit("table-wrap");
function tableInit(block, width, height, cellWidth, cellHeight){
  var charArray = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];
  var parentBlock = document.getElementById(block);

  var viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0)
  var viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0)



  if (width == undefined)  {parentBlock.style.width  = (viewportWidth + "px");}
  if (height == undefined) {parentBlock.style.height = (viewportHeight + "px");}
  if (cellWidth == undefined) {cellWidth = 80;}
  if (cellHeight == undefined) {cellHeight = 20;}

  var rowsInitialAmount = Math.ceil(viewportHeight / cellHeight);
  var columnsInitialAmount = Math.ceil(viewportWidth / cellWidth);

  var tableWrap = document.createElement("DIV");
  tableWrap.id = "super-table-wrap";
  parentBlock.appendChild(tableWrap);

  var creatingGridTableDom = (function() {

    var table = document.createElement("TABLE");
    table.id = "super-table-grid";
    parentBlock.appendChild(table);

    for (var i = 0; i < rowsInitialAmount + 10; i++) {
      var row = document.createElement("TR");
      row.id = "row-grid" + i;
      table.appendChild(row);
      for (var j = 0; j < columnsInitialAmount + 10; j++) {
        var cell = document.createElement("TD");
        cell.id = "cell-grid" + charArray[j] + i;
        cell.width = "80px";
        row.appendChild(cell);
      }
    }
    tableWrap.onscroll = function() {
      console.log(tableWrap.scrollTop)
      parentBlock.scrollTop = tableWrap.scrollTop;
    }
  })();

  var creatingTableDom = (function() {

    var table = document.createElement("TABLE");
    table.id = "super-table";
    tableWrap.appendChild(table);

    for (var i = 0; i < rowsInitialAmount + 10; i++) {
      var row = document.createElement("TR");
      row.id = "row" + i;
      table.appendChild(row);
      for (var j = 0; j < columnsInitialAmount + 10; j++) {
        var cell = document.createElement("TD");
        cell.id = "cell" + charArray[j] + i;
        cell.width = "80px";
        row.appendChild(cell);
      }
    }

  })();

  console.log(rowsInitialAmount);
  console.log(columnsInitialAmount);
};