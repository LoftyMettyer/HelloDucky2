//
// Resizable Table Columns.
//  version: 1.0
//
// (c) 2006, bz
//
// 25.12.2006:  first working prototype
// 26.12.2006:  now works in IE as well but not in Opera 
// 27.12.2006:  changed initialization, now just make class='resizable' in table and load script
//

function preventEvent(e) {
  var ev = e || window.event;
  if (ev.preventDefault) ev.preventDefault();
  else ev.returnValue = false;
  if (ev.stopPropagation)
    ev.stopPropagation();
  return false;
}

function getStyle(x, styleProp) {
  if (x.currentStyle)
    var y = x.currentStyle[styleProp];
  else if (window.getComputedStyle)
    var y = document.defaultView.getComputedStyle(x, null).getPropertyValue(styleProp);
  return y;
}

function getWidth(x) {
  if (x.currentStyle)
  // in IE
    var y = x.clientWidth - parseInt(x.currentStyle["paddingLeft"]) - parseInt(x.currentStyle["paddingRight"]);
  // for IE5: var y = x.offsetWidth;
  else if (window.getComputedStyle)
  // in Gecko
    var y = document.defaultView.getComputedStyle(x, null).getPropertyValue("width");
  return y || 0;
}

function setCookie(name, value, expires, path, domain, secure) {
  document.cookie = name + "=" + escape(value) +
		((expires) ? "; expires=" + expires : "") +
		((path) ? "; path=" + path : "") +
		((domain) ? "; domain=" + domain : "") +
		((secure) ? "; secure" : "");
}

function getCookie(name) {
  var cookie = " " + document.cookie;
  var search = " " + name + "=";
  var setStr = null;
  var offset = 0;
  var end = 0;
  if (cookie.length > 0) {
    offset = cookie.indexOf(search);
    if (offset != -1) {
      offset += search.length;
      end = cookie.indexOf(";", offset)
      if (end == -1) {
        end = cookie.length;
      }
      setStr = unescape(cookie.substring(offset, end));
    }
  }
  return (setStr);
}
// main class prototype
function ColumnResize(table) {

  if (table.tagName != 'TABLE') return;

  this.id = table.id;

  // ============================================================
  // private data
  var self = this;

  var dragColumns = table.rows[0].cells; // first row columns, used for changing of width
  if (!dragColumns) return; // return if no table exists or no one row exists

  var GridID = table.id.replace("Header", "Grid");
  if (!eval(document.getElementById(GridID))) { return; }
  var dragColumns2 = document.getElementById(GridID).rows[0].cells;

  var dragColumnNo; // current dragging column
  var dragX;        // last event X mouse coordinate

  var saveOnmouseup;   // save document onmouseup event handler
  var saveOnmousemove; // save document onmousemove event handler
  var saveBodyCursor;  // save body cursor property

  // ============================================================
  // methods

  // ============================================================
  // do changes columns widths
  // returns true if success and false otherwise
  this.changeColumnWidth = function(no, w) {
    if (!dragColumns) return false;

    if (no < 0) return false;
    if (dragColumns.length < no) return false;
    
    //Minimum width to the left of the drag point
    if (parseInt(dragColumns[no].style.width) < 50) {
      w1 = 50 - parseInt(dragColumns[no].style.width);
      dragColumns[no].style.width = '50px';
      dragColumns2[no].style.width = '50px';
      if (dragColumns[no + 1]) {
        dragColumns[no + 1].style.width = parseInt(dragColumns[no + 1].style.width) - w1 + 'px';
        dragColumns2[no + 1].style.width = parseInt(dragColumns2[no + 1].style.width) - w1 + 'px';
      }
      return true;
    }
    
    //Minimum width to the right of the drag point
    if (parseInt(dragColumns[no + 1].style.width) < 50) {
      w1 = 50 - parseInt(dragColumns[no + 1].style.width);
      dragColumns[no + 1].style.width = '50px';
      dragColumns2[no + 1].style.width = '50px';
      if (dragColumns[no]) {
        dragColumns[no].style.width = parseInt(dragColumns[no].style.width) - w1 + 'px';
        dragColumns2[no].style.width = parseInt(dragColumns2[no].style.width) - w1 + 'px';
      }
      return true;
    }

    if (parseInt(dragColumns[no].style.width) <= -w) return false;
    if (dragColumns[no + 1] && parseInt(dragColumns[no + 1].style.width) <= w) return false;

    dragColumns[no].style.width = parseInt(dragColumns[no].style.width) + w + 'px';
    dragColumns2[no].style.width = parseInt(dragColumns2[no].style.width) + w + 'px';

    if (dragColumns[no + 1]) {
      dragColumns[no + 1].style.width = parseInt(dragColumns[no + 1].style.width) - w + 'px';
      dragColumns2[no + 1].style.width = parseInt(dragColumns2[no + 1].style.width) - w + 'px';
      return true;
    }
  }

  // ============================================================
  // do drag column width
  this.columnDrag = function(e) {
    var e = e || window.event;
    var X = e.clientX || e.pageX;
    if (!self.changeColumnWidth(dragColumnNo, X - dragX)) {
      // stop drag!      
      self.stopColumnDrag(e);
    }

    dragX = X;
    // prevent other event handling
    preventEvent(e);
    return false;
  }

  // ============================================================
  // stops column dragging
  this.stopColumnDrag = function(e) {
    var e = e || window.event;
    if (!dragColumns) return;

    // restore handlers & cursor
    document.onmouseup = saveOnmouseup;
    document.onmousemove = saveOnmousemove;
    document.body.style.cursor = saveBodyCursor;

    // remember columns widths in cookies for server side
    var colWidth = '';
    var separator = '';
    for (var i = 0; i < dragColumns.length; i++) {
      colWidth += separator + parseInt(getWidth(dragColumns[i]));
      separator = '+';
    }
    var expire = new Date();
    expire.setDate(expire.getDate() + 365); // year
    document.cookie = self.id + '-width=' + colWidth +
			'; expires=' + expire.toGMTString();
    preventEvent(e);
  }

  // ============================================================
  // init data and start dragging
  this.startColumnDrag = function(e) {
    var e = e || window.event;

    // if not first button was clicked
    //if (e.button != 0) return;

    // remember dragging object
    dragColumnNo = (e.target || e.srcElement).parentNode.parentNode.cellIndex;
    dragX = e.clientX || e.pageX;

    // set up current columns widths in their particular attributes
    // do it in two steps to avoid jumps on page!
    var colWidth = new Array();
    for (var i = 0; i < dragColumns.length; i++)
      colWidth[i] = parseInt(getWidth(dragColumns[i]));
    for (var i = 0; i < dragColumns.length; i++) {
      dragColumns[i].width = ""; // for sure
      dragColumns[i].style.width = colWidth[i] + "px";
      dragColumns2[i].width = ""; // tie up the grid column widths too.
      dragColumns2[i].style.width = colWidth[i] + "px";
    }

    saveOnmouseup = document.onmouseup;
    document.onmouseup = self.stopColumnDrag;

    saveBodyCursor = document.body.style.cursor;
    document.body.style.cursor = 'w-resize';

    // fire!
    saveOnmousemove = document.onmousemove;
    document.onmousemove = self.columnDrag;

    preventEvent(e);
  }

  // prepare table header to be draggable
  // it runs during class creation
  for (var i = 0; i < dragColumns.length; i++) {
    dragColumns[i].innerHTML = "<div style='position:relative;text-overflow:ellipsis;height:100%;width:98%'>" +
			"<div onclick='event.cancelBubble=true;' style='" +
			"position:absolute;height:100%;width:5px;margin-right:-5px;" +
			"left:100%;top:0px;cursor:w-resize;z-index:10;'>" +
			"</div>" +
			dragColumns[i].innerHTML + 
			"</div>";
    // BUGBUG: calculate real border width instead of 5px!!!
    dragColumns[i].firstChild.firstChild.onmousedown = this.startColumnDrag;
  }
}

// select all tables and make resizable those that have 'resizable' class
var resizableTables = new Array();
function ResizableColumns() {

  var tables = document.getElementsByTagName('table');
  for (var i = 0; tables.item(i); i++) {
    if (tables[i].className.match(/resizable/)) {

      // generate id
      if (!tables[i].id) tables[i].id = 'table' + (i + 1);
      // make table resizable
      resizableTables[resizableTables.length] = new ColumnResize(tables[i]);
    }
  }
  //	alert(resizableTables.length + ' tables was added.');
}
// init tables
/*
if (document.addEventListener)
document.addEventListener("onload", ResizableColumns, false);
else if (window.attachEvent)
window.attachEvent("onload", ResizableColumns);
*/
try {
  window.addEventListener('load', ResizableColumns, false);
} catch (e) {
  window.onload = ResizableColumns;
}


//document.body.onload = ResizableColumns;

//============================================================
//
// Usage. In your html code just include the follow:
//
//============================================================
// <table id='objectId'>
// ...
// </table>
// < script >
// var xxx = new ColumnDrag( 'objectId' );
// < / script >
//============================================================
//
// NB! spaces was used to prevent browser interpret it!
//
//============================================================

/*
* A very simple script to filter a table according to search criteria
*
* http://leparlement.org/filterTable
* See also http://www.vonloesch.de/node/23
*/
function filterTable(term, tableID) {
  
  //fault HRPRO-2289
  var str = term.value;
  if (str == "filter page...") return false;

  var table = document.getElementById(tableID);

  dehighlight(table);
  var terms = term.value.toLowerCase().split(" ");

  for (var r = 0; r < table.rows.length; r++) {
    var display = '';
    for (var i = 0; i < terms.length; i++) {
      
      //NPG20120202 Fault HRPRO-1923
      //Last [hidden] column in the row stores the ID (record selectors only). So trim them off before filtering

      var lowerHTML = table.rows[r].innerHTML.toLowerCase();

      if (lowerHTML.indexOf("display: none") < 0) {
        var dataRow = table.rows[r].innerHTML.substring(0, lowerHTML.indexOf("display:none")) + '>';
      }
      else {
        var dataRow = table.rows[r].innerHTML.substring(0, lowerHTML.indexOf("display: none")) + '>';
      }

      //Fault 1833 - remove & from search text.
      dataRow = dataRow.replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ');

      //if (table.rows[r].innerHTML.replace(/<[^>]+>/g, "|").toLowerCase()
      if (dataRow.replace(/<[^>]+>/g, "|").toLowerCase()
				.indexOf(terms[i]) < 0) {
        display = 'none';
      } else {
        if (terms[i].length) highlight(terms[i], table.rows[r]);
      }
      table.rows[r].style.display = display;
    }
  }
}


/*
* Transform back each
* <span>preText <span class="highlighted">term</span> postText</span>
* into its original
* preText term postText
*/
function dehighlight(container) {

  for (var i = 0; i < container.childNodes.length; i++) {   
    var node = container.childNodes[i];

    if (node.attributes && node.attributes['class']
			&& node.attributes['class'].value == 'highlighted') {
      node.parentNode.parentNode.replaceChild(
        document.createTextNode(
          node.parentNode.innerHTML.replace(/<[^>]+>/g, "").replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')),
          node.parentNode);
      // Stop here and process next parent
      return;
    } else if (node.nodeType != 3) {
      // Keep going onto other elements
      dehighlight(node);
    }
  }
}

/*
* Create a
* <span>preText <span class="highlighted">term</span> postText</span>
* around each search term
*/
function highlight(term, container) {
  for (var i = 0; i < container.childNodes.length; i++) {
    var node = container.childNodes[i];

    if (node.nodeType == 3) {
      // Text node
      var data = node.data;
      var data_low = data.toLowerCase();
      if (data_low.indexOf(term) >= 0) {
        //term found!
        var new_node = document.createElement('span');

        node.parentNode.replaceChild(new_node, node);

        var result;
        while ((result = data_low.indexOf(term)) != -1) {
          new_node.appendChild(document.createTextNode(
								data.substr(0, result)));
          new_node.appendChild(create_node(
								document.createTextNode(data.substr(
										result, term.length))));
          data = data.substr(result + term.length);
          data_low = data_low.substr(result + term.length);
        }
        new_node.appendChild(document.createTextNode(data));
      }
    } else {
      // Keep going onto other elements
      highlight(term, node);
    }
  }  
}

function create_node(child) {
  var node = document.createElement('span');
  node.setAttribute('class', 'highlighted');
  node.attributes['class'].value = 'highlighted';
  node.appendChild(child);
  return node;
}



/*
* Here is the code used to set a filter on all filterable elements, usually I
* use the behaviour.js library which does that just fine
*/
tables = document.getElementsByTagName('table');
for (var t = 0; t < tables.length; t++) {
  element = tables[t];

  if (element.attributes['class']
		&& element.attributes['class'].value == 'resizable') {

    /* Here is dynamically created a form */
    var form = document.createElement('form');
    form.setAttribute('class', 'filter');
    // For ie...
    form.attributes['class'].value = 'filter';
    var input = document.createElement('input');
    input.onkeyup = function() {
      filterTable(input, element);
    }
    form.appendChild(input);
    element.parentNode.insertBefore(form, element);
  }
}

