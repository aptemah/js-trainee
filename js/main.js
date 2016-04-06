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

  rowsInitialAmount = Math.ceil(viewportHeight / cellHeight);
  columnsInitialAmount = Math.ceil(viewportWidth / cellWidth);

  var creatingTableDom = (function(){
    var table = document.createElement("TABLE");
    table.id = "super-table";
    parentBlock.appendChild(table);

    for (var i = 0; i < rowsInitialAmount; i++) {
      var row = document.createElement("TR");
      row.id = "row" + i;
      table.appendChild(row);
      for (var j = 0; j < columnsInitialAmount; j++) {
        var cell = document.createElement("TD");
        cell.id = "cell_" + charArray[j] + i;
        row.appendChild(cell);
        addClickEvent(cell);
      }
    }

    function addClickEvent(cell){
      cell.onclick = function(){
        var input = document.createElement("INPUT");
        input.id = "inputOf_" + this.id;
        this.appendChild(input);
        input.focus();
        var that = this;
        input.onblur = function(){
          that.removeChild(this);
        };
      };
    };

  })();

};