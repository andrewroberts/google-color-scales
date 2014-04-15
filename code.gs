//234567890123456789012345678901234567890123456789012345678901234567890123456789

/*

Colour Scales
=============

This script colours in cells in a selection of cells containing 
numbers by the relative value. This is similar to Color Scales in 
Excel.

Version 0.2
-----------
- Added a dialog box to choose the colours rather than hard-coding 
them (they are now stored in UserProperties).
- Added single call (setBackgrounds()) to set backgrounds plus other 
optimisations.

Version 0.1 
-----------
- First Version.

*/

function onOpen() {
  
  var subMenus = [{name:"Colour Scale Selection", functionName: "cs_colourScale"},
                  {name:"Choose Colours", functionName: "cs_chooseColourUI"}];
  
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Colour Scale", subMenus);
  
}

/*

cs_number()
===========

The number class.

@class cs_number

*/

cs_number = function() {
  
  this.value; 
  this.row;
  this.column;
  
}

/* 

cs_colour()
===========

The colour class.

@class cs_colour

*/

cs_colour = function(red, green, blue) {
  
  this.red = red;
  this.green = green;
  this.blue = blue;
  
}

/*

cs_colourScale()
================

Set the cell background colour depending on its numeric value.

@params none

@returns none

*/

function cs_colourScale() {

  var range = SpreadsheetApp.getActiveRange();
  var cellValues = range.getValues();
  var numRows = range.getNumRows();
  var numColumns = range.getNumColumns();
  
  // TODO - Check that colours have been defined.

  // First sweep
  // -----------
  //
  // Do a first sweep to get the highest and lowest number 
  // in the range.
  //

  var highest = new cs_number();
  var lowest = new cs_number(); 
  
  for (var i = 0; i < numRows; i++) {
    
    for (var j = 0; j < numColumns; j++) {
      
      var tempVal = cellValues[i][j];
      
      if (typeof tempVal !== 'number') {
        
        // Ignore this value as it's not a number.
        continue;
        
      }
      
      if (tempVal > highest.value || highest.value === undefined) {
        
        // Record the highest number found so far.
        highest = {'value': tempVal, 'row': i, 'column': j};
        
      }
      
      if (tempVal < lowest.value || lowest.value === undefined) {
        
        // Record the lowest number found so far.
        lowest = {'value': tempVal, 'row': i, 'column': j};
                
      }      
    } // for j     
  } // for i

  // Check the highest and lowest values.
  if (lowest.value === undefined && highest.value !== undefined) {
    
    lowest.value = highest.value;
    
  }
  
  if (highest.value === undefined && lowest.value !== undefined) {
    
    highest.value = lowest.value;
    
  }
      
  if (lowest.value === undefined && highest.value === undefined) {
    
    Browser.msgBox("No numbers found in selection");
    return;
    
  }

  //
  // Second Sweep
  // ------------
  //
  // Do a second sweep through the range and set the cell colour.
  //
  
  // Get the user defined low and high colours.
  var lowRgb = {red: Number(UserProperties.getProperty('lowR')),
                green: Number(UserProperties.getProperty('lowG')),
                blue: Number(UserProperties.getProperty('lowB'))};
                
  var highRgb = {red: Number(UserProperties.getProperty('highR')),
                 green: Number(UserProperties.getProperty('highG')),
                 blue: Number(UserProperties.getProperty('highB'))};
  
  var cellColours = [];
  
  for (i = 0; i < numRows; i++) {
    
    // Create an array of colours for this row.
    var rowColours = [];
    
    for (j = 0; j < numColumns; j++) {
      
      // Set the colour depending on the value compared to the 
      // high and low values.
          
      // TODO - Error check getCellColour_
      
      // Store the colours for now to write later.
      rowColours.push(getCellColour_(cellValues[i][j], 
                                     lowest.value, 
                                     highest.value,
                                     lowRgb,
                                     highRgb));
            
    } // for j
    
    // Add the next row of colours to the main array.
    cellColours.push(rowColours);
    
  } // for i
  
  range.setBackgrounds(cellColours);
  
} // cs_colourScale()

/*

getCellColour_()
================

Work out what colour this cell should be.

@param {Number} value The numeric value in the present cell.
@param {Number} lowValue The lowest value found in the selected cells.  
@param {Number} highValue The highest value found in the selected cells.
@param {cs_colour} lowRgb Object describing the lowest colour.
@param {cs_colour} highRgb Object describing the highest colour.

@returns {String} 7 character hex colour string.

*/

function getCellColour_(value, lowValue, highValue, lowRgb, highRgb) {

  // Check the parameters.
  
  if (typeof value != 'number') {
    
    // Not a number so just return the low colour value.
    return rgbToHex_(lowRgb.red,lowRgb.green,lowRgb.blue);
    
  }

  if (typeof lowValue !== 'number' || typeof highValue !== 'number') {
    
    throw Error("NaN passed to getCellColour_()");
    
  }
  
  if (typeof lowRgb != 'object' || typeof highRgb != 'object') {
    
    throw Error("Colours passed to getCellColour_() not objects");
    
  }
    
  // The increment in a colour is in the same ratio as the number value
  // to the maximum number difference.

  var valueDiff = highValue - lowValue;
                   
  if (valueDiff === 0) {
    
    // There is no difference in colours so just return the 
    // low colour value.
    return rgbToHex_(lowRgb.red,lowRgb.green,lowRgb.blue)
    
  }
  
  var valueRatio = (value - lowValue)/valueDiff;

  var redDiff = Math.abs(highRgb.red - lowRgb.red);  
  var inc = Math.floor(valueRatio * redDiff);
  var red = (lowRgb.red < highRgb.red) ? (lowRgb.red + inc) : (lowRgb.red - inc);
  
  var greenDiff = Math.abs(highRgb.green - lowRgb.green);
  inc = Math.floor(valueRatio * greenDiff);
  var green = (lowRgb.green < highRgb.green) ? (lowRgb.green + inc) : (lowRgb.green - inc);

  var blueDiff = Math.abs(highRgb.blue - lowRgb.blue);  
  inc = Math.floor(valueRatio * blueDiff);
  var blue = (lowRgb.blue < highRgb.blue) ? (lowRgb.blue + inc) : (lowRgb.blue - inc);
  
  return rgbToHex_(red,green,blue);
  
} // getCellColour_()

/*

cs_chooseColourUI_
==================

UI to allow the user to choose the colour.

@params none

@returns none

*/

function cs_chooseColourUI()
{

  var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  var app = UiApp.createApplication()
                 .setTitle("Choose the low and high colour:")
                 .setHeight('100')
                 .setWidth('400');
  
  // Create the form panel to hold all content in the UI for this step
  // (this can only contain one widget).
  var formPanel = app.createFormPanel()
                     .setStyleAttribute('border-spacing', '10px');
  
  // Create a flow panel to contain everything else.
  var flowPanel = app.createFlowPanel();
  
  // Create the grid to hold the questions.
  var grid = app.createGrid(2, 4);
  
  grid.setWidget(0, 0, app.createLabel('Low Colour RGB (e.g. 255,0,255):'));
  grid.setWidget(0, 1, app.createTextBox().setName('lowR').setWidth('30').setValue(UserProperties.getProperty('lowR')));
  grid.setWidget(0, 2, app.createTextBox().setName('lowG').setWidth('30').setValue(UserProperties.getProperty('lowG')));
  grid.setWidget(0, 3, app.createTextBox().setName('lowB').setWidth('30').setValue(UserProperties.getProperty('lowB')));  
  grid.setWidget(1, 0, app.createLabel('High Colour RGB (e.g. 255,0,255):'));
  grid.setWidget(1, 1, app.createTextBox().setName('highR').setWidth('30').setValue(UserProperties.getProperty('highR')));
  grid.setWidget(1, 2, app.createTextBox().setName('highG').setWidth('30').setValue(UserProperties.getProperty('highG')));
  grid.setWidget(1, 3, app.createTextBox().setName('highB').setWidth('30').setValue(UserProperties.getProperty('highB')));
  
  flowPanel.add(grid);
  flowPanel.add(app.createSubmitButton("Submit"));
  formPanel.add(flowPanel);
  app.add(formPanel);

  ss.show(app);
         
}

/*

doPost()
========

Event handler for POST event from colour choosing UI.

@params {Object} eventInfo Post info from UI.

@returns {Object} app App displaying UI.

*/
  
function doPost(eventInfo) {
  
  // TODO - Error checking on RGB values.

  // Set user properties from the colour selection form.  
  UserProperties.setProperty('lowR', eventInfo.parameter.lowR);
  UserProperties.setProperty('lowG', eventInfo.parameter.lowG);
  UserProperties.setProperty('lowB', eventInfo.parameter.lowB);

  UserProperties.setProperty('highR', eventInfo.parameter.highR);
  UserProperties.setProperty('highG', eventInfo.parameter.highG);
  UserProperties.setProperty('highB', eventInfo.parameter.highB);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clean up - get the UiInstance object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();

  // The following line is REQUIRED for the widget to actually close.
  return app;

}

/*
  
rgbToHex_
=========

Convert RGB to hex value. Thanks to Tim Down: 

  http://stackoverflow.com/questions/5623838/rgb-to-hex-and-hex-to-rgb

@params {Number} red 
@params {Number} green
@params {Number} blue

@returns {String} hex string

*/

function rgbToHex_(red, green, blue) {
  
  var rgbColourToHex_ = function(colourNum) {
  
    // Check it's less that 255 (hex FF).
    if (colourNum > 255) return '00';
    
    // Round the number down and convert it to a hex string.
    var hex = Math.floor(colourNum).toString(16);
    
    // Pad to two characters.
    return hex.length == 1 ? '0' + hex : hex;
    
  };
  
  var hexString = '#' + 
                  rgbColourToHex_(red) + 
                  rgbColourToHex_(green) + 
                  rgbColourToHex_(blue);
  
  return hexString;

} // rgbToHex_()
