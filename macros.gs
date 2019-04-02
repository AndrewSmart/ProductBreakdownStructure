/**
 * Macros (javascript) intended for use on a Product Breakdown Structure (PBS)
 *  spreadsheet, to automate its creation and update. Intended for use on
 *  a Google Spreadsheet (not MS Excel).
 *
 * Macros:
 * update_pbs_numbers- Classifies assemblies/parts and assigns a PBS#.
 * populate_tabs- Creates a BOM split into 3 tabs using data from the PBS tab.
 *
 * All other functions are only used internally. Import only the above two
 *  macros into the Google Spreadsheet, of proper PBS form (see example).
 *
 * A brief description of the PBS tab this code is indexed for is as follows,
 *  first two rows are header, data starts on row 3, and column are:
 *  PBS#, CAD ID, Name, X, Type, Build, Quantity
 *  (X is column not used by these macros)
 *  PBS# has a form of ####-## (4 numbers for hierarchy depth, followed
 *  by a zero padded part number). If your project has deeper depth, add another #.
 *  Name has a ' ' and '-' padded prefix to indicate depth.
 *  Type is an enum including 'A - Assembly', 'D - Designed Part', 'H - Hardware'
 *  Build is an enum including '3 - 3D Print', 'O - Off the Shelf', 'C - Carbon Fiber'
 *
 * Uses the Google Apps Script API:
 * https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * https://developers.google.com/apps-script/reference/spreadsheet/range
 *
 * Maintained here: https://github.com/AndrewSmart/ProductBreakdownStructure
 *
 * Copyright © 2018 Andrew Smart
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along
 * with this program; if not, write to the Free Software Foundation, Inc.,
 * 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 */

/*
 * Autopopulates the tabs from the PBS tab.
 * Useful to know how many parts to print, buy, etc. from data in PBS spreadsheet.
 * Would be nice to link spreadsheet up to a CAD system, like I know FreeCAD can do spreadsheets.
 */
function populate_tabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pbsSheet = ss.getSheetByName('PBS');
  var printsSheet = ss.getSheetByName('BOM-3D Prints');
  var hardwareSheet = ss.getSheetByName('BOM-Hardware');
  var carbonFiberSheet = ss.getSheetByName('BOM-CarbonFiber');
  //var pbsData = pbsSheet.getDataRange().getValues();
  var autoPopulatedSheets = [printsSheet, hardwareSheet, carbonFiberSheet];
  //var log = pbsSheet.getRange(3, 1); log.setValue(log.getValue() + " Hi");

  // Clear the autopopulated sheet; delete all but first data row, so style not lost
  autoPopulatedSheets.forEach(function(sht) {
    if(sht.getMaxRows() > 3)
      sht.deleteRows(4, sht.getMaxRows()-3)
    sht.getRange(3, 1, 1, sht.getMaxColumns()).clear();
  });

  /* LINEAR TRAVERSAL OF PBS TAB, has bug in that quantity should be multiplied by parent assemblies, impossible to do linearly.
  // Put Hardware into Hardware tab, and 3D Prints into 3D Print tab
  for(var i = 4; i < pbsData.length; ++i) {
    var buildType = pbsData[i][5].toString();
    var pbsHyperlink = "=HYPERLINK(\"#gid=0&range=A" + (i+1) + "\",\"" + pbsData[i][0] + "\")";
    if(buildType === "3 - 3D PRINT") {
      printsSheet.appendRow([pbsHyperlink, pbsData[i][1], pbsData[i][2].toString().replace(/^[- ]+/,''), pbsData[i][6], pbsData[i][7], pbsData[i][9]]);
    } else if(buildType === "O - OFF THE SHELF") {
      hardwareSheet.appendRow([pbsHyperlink, pbsData[i][1], pbsData[i][2].toString().replace(/^[- ]+/,''), pbsData[i][6], pbsData[i][7], pbsData[i][9]]);
    } else if(buildType === "C - CARBON FIBER") {
      carbonFiberSheet.appendRow([pbsHyperlink, pbsData[i][1], pbsData[i][2].toString().replace(/^[- ]+/,''), pbsData[i][6], pbsData[i][7], pbsData[i][9]]);
    }
  }
  */
  // Tree Traversal of PBS Tab:
  var rootNode = pbsTree_makeTree();
  validate_pbsTree(rootNode, pbsSheet); //Set assemblies if blank, set leaf nodes as printed parts by default
  var pbsData = pbsSheet.getRange(3, 1, pbsSheet.getMaxRows()-2, pbsSheet.getMaxColumns()).getValues();
  populate_tabs_from_pbsTree(rootNode, 1, pbsData);

  // Remove prior first row and sort:
  autoPopulatedSheets.forEach(function(sht) {
    //if(sht.getMaxRows() > 3)
    //  sht.deleteRow(3);
    sht.getRange(3, 1, sht.getMaxRows()-2, sht.getMaxColumns()).sort([3,6]);
  });
  group_sheet(printsSheet);
  group_sheet(hardwareSheet);
};

/** Function assumes sheet already sorted by matchColumn.*/
function group_sheet(sht) {
  var pbsNames = sht.getRange(3, 1, sht.getMaxRows()-2, sht.getMaxColumns()).getValues();
  // Search from the bottom, so that we don't have to deal with indicies into the sheet changing.
  for(var i = pbsNames.length - 1; i > 0; --i) {
    var j = i - 1;
    var quantity = pbsNames[i][3] > 1 ? pbsNames[i][3] : 1;
    //Regex unreliable, so using for loop to find # suffix.
    var k = pbsNames[i][2].length - 1;
    for(; k >= 0; --k) {
      var ch = pbsNames[i][2].charAt(k);
      if(ch < "0" || ch > "9")
        break;
    }
    outerBase = pbsNames[i][2].substring(0, k+1);
    //log.setValue(log.getValue() + " oB:" + outerBase);
    for(; j >= 0; --j) { // Search for matches above i, keep looping until difference
      if(pbsNames[i][1] != pbsNames[j][1]) { //Check if CAD IDs not the same first
        j++; //increment j for the mismatch
        break;
      }
      var k = pbsNames[j][2].length - 1;
      for(; k >= 0; --k) {
        var ch = pbsNames[j][2].charAt(k);
        if(ch < "0" || ch > "9")
          break;
      }
      innerBase = pbsNames[j][2].substring(0, k+1);
      //log.setValue(log.getValue() + " iB:" + innerBase + '\n');
      if(innerBase != outerBase) {
        j++; //increment j for the mismatch.
        break;
      }
      quantity += pbsNames[j][3] > 1 ? pbsNames[j][3] : 1;
    }
    // Check for row(s) matching
    if(j != i) {
      // Make new row header, with quantity.
      sht.insertRowBefore(j+3);
      var newRow = sht.getRange(j+3, 1, 1, sht.getMaxColumns());
      newRow.setValues([["","","TOTAL INSTANCES--"+pbsNames[j][2]+"--",quantity,pbsNames[j][4],pbsNames[j][5]]]);
      newRow.setBackground('#88aaaa');
      sht.getRange(j+4, 1, i-j+1, 1).shiftRowGroupDepth(1);
      sht.getRowGroup(j+4, 1).collapse();
      i = j; // Resume search above matching lines.
    }
  }
}

/*
 * Validates PBS tab.
 * Makes sure assemblies have vacant 'BUILD'.
 * Makes sure leaf nodes are a part/hardware.
 * Makes sure leaf nodes have a populated 'BUILD'.
 */
function validate_pbsTree(pNode, pbsSheet) {
  for(var j = 0; j < pNode.child.length; ++j) {
    var pChild = pNode.child[j];
    var iChild = pChild.data;
    var partType = pbsSheet.getRange(3+iChild, 5, 1, 1);
    if(0 == pChild.child.length) { //leaf node
      var buildType = pbsSheet.getRange(3+iChild, 6, 1, 1);
      if(buildType.getValue() === "") { //For Leaf node, default to 3D Print, and set on PBS Tab
        if(partType.getValue() === "") {
          partType.setValue("D - DESIGNED PART");
          buildType.setValue("3 - 3D PRINT");
        }
      }
    } else { //assembly node (has children)
      if(partType.getValue() === "") { //Set type to assembly if not set
        partType.setValue("A - ASSEMBLY");
      }
      validate_pbsTree(pChild, pbsSheet); //Handle children
    }
  }
};

function setCharAt(str,index,chr) {
  if(index > str.length-1) return str;
  return str.substr(0,index) + chr + str.substr(index+1);
}

var Node = function(data) {
  this.data = data;
  this.depth = null;
  this.pbs = null;
  this.parent = null;
  this.child = [];
};

/*
 * Builds a tree from the names.
 * Leaf nodes are either parts or hardware.
 * Non-leaf nodes are assemblies.
 * Construction of a tree from names will allow both validation, and autogeneration of the PBS tab from CAD data.
 * API reading CAD data hierarchy should write names with appropriate prefix '-', '  ', etcetera. This function
 * interprets said name output of the CAD API.
 */
function update_pbs_numbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pbsSheet = ss.getSheetByName('PBS');

  var rootNode = pbsTree_makeTree();
  pbsSheet.getRange(3, 1).setValue(rootNode.pbs = "0000-00");
  write_pbsnumber_to_sheet(pbsSheet, rootNode);
};

/** Makes the pbs tree data structure, which can be used to do useful things.*/
function pbsTree_makeTree() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pbsSheet = ss.getSheetByName('PBS');
  var pbsNames = pbsSheet.getRange(3, 3, pbsSheet.getMaxRows()-2, 1).getValues();

  //Build the PBS tree, of nodes, from the name prefixes, we then use the hierarchy to write the PBS#s.
  //Node is an index into pbsNames
  var rootNode = new Node(0);
  //Put root at -1 to offset main children, which have indentation of 0 (nameDepth of 0), so that algorithm works.
  rootNode.depth = -1;
  pbsTree_addChildren(pbsNames, rootNode, 1);
  //var log = pbsSheet.getRange(3, 1); log.setValue(log.getValue() + " done");
  return rootNode;
}

/* Recursive function, adds children starting at index to node.*/
function pbsTree_addChildren(/*pbsSheet,*/ pbsNames, node, index) {
  var child = null;
  var i = index;
  //var log = pbsSheet.getRange(3, 1); log.setValue(log.getValue() + " [" + index + "]:" + pbsNames[i][0]);
  for(; i < pbsNames.length; ) { //Loop over possible children of node, and recurse if child of child
    var indexDepth = nameDepth(pbsNames[i][0]);
    //log.setValue(log.getValue() + " [nD:" + node.depth + " iD:" + indexDepth + "]");
    if(node.depth + 1 == indexDepth) {
      // New child of node.
      child = new Node(i);
      child.parent = node;
      child.depth = indexDepth;
      node.child.push(child);
      ++i;
    } else if(node.depth + 1 < indexDepth) {
      // Now a child of child, recurse:
      i = pbsTree_addChildren(pbsNames, child, i);
    } else if(node.depth + 1 > indexDepth) {
      // No more children, return to parent stopping index so parent can pick up remaining children here.
      //log.setValue(log.getValue() + " [return row:" + (i+3) + " nD:" + node.depth + " iD:" + indexDepth + "]\n");
      return i;
    }
  }
  return i; //Reached end of pbsNames (end of list), return up hierachy
};

/* Gets the hierarchy depth of a name. Root is 0, "-Assembly" is 1, "  Part" is 2.*/
function nameDepth(name) {
  var i = 0;
  for(; i < name.length; ++i) {
    if(name.charAt(i) != ' ' && name.charAt(i) != '-') {
      break;
    }
  }
  return i;
}

/* A debug function, writes the hierarchy depth of each part/assembly in column 11.*/
function write_nodeDepth_to_sheet(pbsSheet, node) {
  for(var i = 0; i < node.child.length; ++i) {
    //var log = pbsSheet.getRange(3, 1);
    //log.setValue(log.getValue() + node.child.length);
    //var pbsName = pbsSheet.getRange(node.child[i].data + 3, 3).getValue();
    pbsSheet.getRange(node.child[i].data + 3, 11).setValue(node.child[i].depth);//nameDepth(pbsName));
    write_nodeDepth_to_sheet(pbsSheet, node.child[i]);
  }
};

/* Writes PBS #s to sheet, using the tree. Doesn't write root node. Convieniently no assembly in hierarchy exceeds 9 peers.*/
function write_pbsnumber_to_sheet(pbsSheet, pNode) {
  var partsWithChildren = 0;
  for(var i = 0; i < pNode.child.length; ++i) {
    if(0 == pNode.child[i].child.length)
      pNode.child[i].pbs = pNode.pbs.substr(0,5) + ("00" + (i+1)).slice(-2); // leaf node
    else {
      ++partsWithChildren;
      pNode.child[i].pbs = setCharAt(pNode.pbs, pNode.child[i].depth, partsWithChildren);
    }
    pbsSheet.getRange(pNode.child[i].data + 3, 1).setValue(pNode.child[i].pbs);
    write_pbsnumber_to_sheet(pbsSheet, pNode.child[i]);
  }
};

/* Populate tabs from PBS Tree.
 * @pNode pointer into the tree data structure
 * @runningQuantity
 * @pbsData contains the PBS tab data minus the header rows, reduces sheet API calls using it
 */
function populate_tabs_from_pbsTree(pNode, runningQuantity, pbsData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pbsSheet = ss.getSheetByName('PBS');
  var printsSheet = ss.getSheetByName('BOM-3D Prints');
  var hardwareSheet = ss.getSheetByName('BOM-Hardware');
  var carbonFiberSheet = ss.getSheetByName('BOM-CarbonFiber');
  for(var j = 0; j < pNode.child.length; ++j) {
    var pChild = pNode.child[j];
    var iChild = pChild.data;
    var quantity = pbsData[iChild][6] > 1 ? pbsData[iChild][6] : 1;
    var isDeprecated = pbsData[iChild][7] === "X - DEPRECATED";
    quantity = runningQuantity * quantity;
    if(!isDeprecated) {
      if(0 == pChild.child.length) { //leaf node
        // Find parent assembly quantities, and multiply each.
        //var log = pbsSheet.getRange(3, 1);
        //log.setValue(log.getValue() + " q:" + quantity);
        var buildType = pbsData[iChild][5].toString();
        var pbsHyperlink = "=HYPERLINK(\"#gid=0&range=A" + (iChild+3) + "\",\"" + pbsData[iChild][0] + "\")";
        if(buildType === "3 - 3D PRINT") {
          printsSheet.appendRow([pbsHyperlink, pbsData[iChild][1], pbsData[iChild][2].toString().replace(/^[- ]+/,''), quantity, pbsData[iChild][7], pbsData[iChild][9]]);
        } else if(buildType === "O - OFF THE SHELF") {
          hardwareSheet.appendRow([pbsHyperlink, pbsData[iChild][1], pbsData[iChild][2].toString().replace(/^[- ]+/,''), quantity, pbsData[iChild][7], pbsData[iChild][9]]);
        } else if(buildType === "C - CARBON FIBER") {
          carbonFiberSheet.appendRow([pbsHyperlink, pbsData[iChild][1], pbsData[iChild][2].toString().replace(/^[- ]+/,''), quantity, pbsData[iChild][7], pbsData[iChild][9]]);
        }
      } else {//Has children:
        populate_tabs_from_pbsTree(pChild, quantity, pbsData); //Handle children
      }
    }
  }
};
