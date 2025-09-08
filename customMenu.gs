function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('GSheetsAdmin')
      .addItem('Create Files', 'menuItem1')
      .addItem('Update File Info and Links', 'menuItem2')
      .addItem('Retrieve Grades', 'menuItem3')
      .addItem('Share Solution file','menuItem4')
      .addSeparator()
      .addSubMenu(ui.createMenu('Modify file permissions')
          .addItem('Grant editor access', 'menu2Item1')
          .addItem('Grant viewer access', 'menu2Item2')
          .addItem('Revoke editor access', 'menu2Item3')
          .addItem('Revoke viewer access', 'menu2Item4')
          .addItem('Revoke sharing permission', 'menu2Item5'))
      .addToUi();
}

function menuItem1() { //Create template files
  if (confirmYN()==1) {
    SpreadsheetApp.getUi().alert('Create them files!');
    createTemplates(1);
    return;
  }
}

function menuItem2() { // Update file info and links, 
  createTemplates(11);
  return;

}

function menuItem3() {  // Retrieve grades or other info 
  createTemplates(5);
  return;
}
function menuItem4() {  // Share solution 
  createTemplates(5);
  return;
}
function menu2Item1() { // Grant editor access
  createTemplates(4);
  return;
}

function menu2Item2() { // Grant viewer access
  createTemplates(7);
  return;
}

function menu2Item3() { // Revoke editor access
  createTemplates(5);
  return;
}

function menu2Item4() { // Revoke viewer access
  createTemplates(8);
  return;
}

function menu2Item5() { // Revoke sharing permission
  createTemplates(3);
  return;
}
