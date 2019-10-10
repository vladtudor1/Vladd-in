/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

 var _doc;
 var lastLookup; 

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

 // The initialize function must be run each time a new page is loaded
 Office.initialize = function (reason) {
  $(document).ready(function(){
    _doc = Office.context.document;
    tryUpdatingSelectedWord();
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);
  });
};

// Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
//   if (asyncResult.status == Office.AsyncResultStatus.Failed) {
//       write('Action failed. Error: ' + asyncResult.error.message);
//   }
//   else {
//       write('Selected data: ' + asyncResult.value);
//   }
// });

// // Function that writes to a div with id='message' on the page.
// function write(message){
//   document.getElementById('message').innerText += message; 
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