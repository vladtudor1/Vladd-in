/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Office.onReady(info => {
//   if (info.host === Office.HostType.Word) {
//     //document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//   }
// });

//  // The initialize function must be run each time a new page is loaded
//  Office.initialize = function (reason) {
//   $(document).ready(function(){
//     _doc = Office.context.document;
//     tryUpdatingSelectedWord();
//     _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);
//     $('#getOOXML').click(function() { getOOXML(); });
//     $('#setOOXML').click(function () { setOOXML(); });

//     loadSampleData();
//   });
// };

// function getOOXML() {
//   var report = document.getElementById("status");

//   while (report.hasChildNodes()) {
//       report.removeChild(report.lastChild);
//   }

//   Office.context.document.getSelectedDataAsync(
//       Office.CoercionType.Ooxml,
//       { valueFormat: Office.ValueFormat.Formatted, filterType: Office.FilterType.All },
//       function (result) {
//           var textArea = document.getElementById("dataOOXML");
//           if (result.status == "succeeded") {
//               currentOOXML = result.value;
              

//               while (textArea.hasChildNodes()) {
//                   textArea.removeChild(textArea.lastChild);
//               }
//               textArea.appendChild(document.createTextNode(currentOOXML));

//               // Tell the user we succeeded
//               report.innerText = "Got It --- Success!!";
//           }
//           else {
//               // This runs if the getSnelectedDataAsync method does not return a success flag
//               currentOOXML = "";
//               report.innerText = result.error.message;
//           }
//       });
// }

// function setOOXML() {
//   var report = document.getElementById("status");
//   while (report.hasChildNodes()) {
//       report.removeChild(report.lastChild);
//   }

//   if (currentOOXML != "") {


//       Office.context.document.setSelectedDataAsync(
//           currentOOXML,
//           { coercionType: Office.CoercionType.Ooxml },
//           function (result) {
//               console.log(result)
//               // Update the report element
//               if (result.status == "succeeded") {
//                   report.innerText = "Set It --- Success!!";
//               }
//               else {
//                   report.innerText = result.error.message;
//                   while (textArea.hasChildNodes()) {
//                       textArea.removeChild(textArea.lastChild);
//                   }
//               }
//           });
//   }
//   else {

//       report.innerText = "There is currently no OOXML to insert!"
//           + " Please select some of your document and click [Get OOXML] first!";
//   }
// }

 var _doc;
 var lastLookup; 
 var currentOOXML = "";

(function () {
  "use strict";

  // The initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      $(document).ready(function () {
        OfficeExtension.config.extendedErrorLogging = true;
        _doc = Office.context.document;
          tryUpdatingSelectedWord();
          _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);

          $("#template-description").text("This sample creates a bookmark at the caret position or around a selection.  Note, there are bugs in Word that prevent anchor bookmarks at the caret position.  To get around those bugs I insert '[delete - necessary to avoid bug in Word 2016 (16.0.6925.1018)]' text after an anchor bookmark.  I'm trying to submit this to Microsoft officially as a bug, along with a suggestion to support bookmarks natively just like they support content controls.  Note, Word can create these and the OPC, OOXML, WordProcessingML is the same so the bug is in the '.insertOoxml(...)' methods.");
          $('#button-text').text("Bookmark!");
          $('#button-desc').text("Places a bookmark at the caret or around the selection");

          // loadSampleData();

          // Add a click event handler for the bookmark button.
          $('#bookmark-button').click(
              placeBookmark);
      });
  };

  // This function inserts sample data, or a pre-written piece of text. Don't use outside of development environment

  // function loadSampleData() {

  //     // Run a batch operation against the Word object model.
  //     Word.run(function (context) {

  //         // Create a proxy object for the document body.
  //         var body = context.document.body;

  //         // Queue a commmand to clear the contents of the body.
  //         body.clear();
  //         // Queue a command to insert text into the end of the Word document body.
  //         body.insertText("This is a sample text inserted in the document",
  //                         Word.InsertLocation.end);

  //         // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  //         return context.sync();
  //     })
  //     .catch(errorHandler);
  // }

  function tryUpdatingSelectedWord() {
  _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    if (selectedText != "") { 
        if (selectedText != lastLookup) { 
            lastLookup = selectedText; 
            $("#headword").text(selectedText); 
    }
  }
}

  function placeBookmark() {

      Word.run(function (context) {

          // Create a bookmark ID - 0 works just fine as Word will generate a new number upon insert which is awesome!
          var bkmkId = 0;

          // Create a bookmark Name - it must be unique or it will overwrite a current one!
          // It must be no longer than 40 characters!
          // Please see the behavior here:
          var bkmkName = "_TOC_MANUAL_" + new Date().getTime();

          // Queue a command to get the current selection and then
          // create a proxy range object with the results.
          var range = context.document.getSelection();

          // Use the extension to insert the bookmark.  It really should be this easy.  See the header of the
          // extension for an idea of the current suggested API for Microsoft.  I only wrote a quick insertBookmark
          // method that doesn't quite fit the signature for what's proposed btw.  I wanted this to be quick and
          // minimal.  Plus, the extension should really extend range, document.body, etc.
          officeWordExtension.insertBookmark(bkmkId, bkmkName, handleSuccess, errorHandler);
      });
  }

  function handleSuccess() {
      showNotification("Success:", "It worked, use Word to check for your bookmark!");
  }

  //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
  function errorHandler(error) {
      // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
      showNotification("Error:", error);
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {

      var report = document.getElementById("status");
       while (report.hasChildNodes()) {
      report.removeChild(report.lastChild);
       }

      $("#notificationHeader").text(header);
      $("#notificationBody").text(content);
      report.innerText = "Congrats, your bookmark has been inserted!" 
//           + " Please select some of your document and click [Get OOXML] first!";
      // messageBanner.showBanner();
      // messageBanner.toggleExpansion();
  }
})();


