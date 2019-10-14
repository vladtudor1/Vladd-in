/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

 var _doc;
 var lastLookup; 
 var sFileName;
 var sOoxml;

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
    $('#insertContentControls').click(insertContentControls);
    $('#refBookmark').click(refBookmark);
    $('#bookmarkSelection').click(bookmarkSelection);
  });
};

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

 // Bookmark the current selection
 function bookmarkSelection() {
  //Get the current selection in order to bookmark
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
      function (resultGet) {
          if (resultGet.status === Office.AsyncResultStatus.Succeeded) {
              selection = resultGet.value;
              //Get the WordOpenXML for creating a bookmark
              $.ajax({
                  type: "GET",
                  url: sFileName,
                  dataType: "xml",
                  success: function (resultOoxml) {
                    Ooxml = resultOoxml;
                    var textNode = 
                        Ooxml.documentElement.getElementsByTagNameNS('*', 't');
                    //Replace the text to be bookmarked with the selection
                    textNode[0].textContent = selection;
                    sOoxml = 
                      new XMLSerializer().serializeToString(Ooxml.documentElement);
                    console.log(sOoxml);
                    //Write the bookmark back to the document
                    Office.context.document.setSelectedDataAsync(sOoxml, 
                       { coercionType: Office.CoercionType.Ooxml },
                        function (result) {
                          if (result.status === "succeeded") {
                              app.showNotification('Success!!');
                          } else {
                              app.showNotification('Error:', result.error.message);
                          }
                        });
                    }
              });
          }
          else {
            app.showNotification('Error:', resultGet.error.message);
          }
      });
}

    //   //Insert a reference to the bookmark at the current selection
    // function refBookmark() {

    //   $.ajax({
    //       type: "GET",
    //       url: sFileName,
    //       dataType: "xml",
    //       success: function (resultOoxml) {
    //         Ooxml = resultOoxml;
    //         var textNode = Ooxml.documentElement.getElementsByTagNameNS('*', 't');
    //         //Replace the selection with the Ref field.
    //         //Since it's not possible to "look up" the bookmark using the
    //         //2013 APIs the text content of the bookmark is not known - 
    //         //the user will have to update the field manually.
    //         textNode[0].textContent = "Press F9 to update the field";
    //         sOoxml = new XMLSerializer().serializeToString(Ooxml.documentElement);
    //         console.log(sOoxml);
    //         //Write the Ref field back to the document
    //         Office.context.document.setSelectedDataAsync(sOoxml, 
    //             { coercionType: Office.CoercionType.Ooxml },
    //             function (result) {
    //                 if (result.status === Office.AsyncResultStatus.Succeeded) {
    //                     app.showNotification('Success!!');
    //                 } else {
    //                     app.showNotification('Error:', result.error.message);
    //                 }
    //             });
    //       }
    //   });
    // }

     async function insertContentControls() {
      await Word.run(async (context) => {
          // Create a proxy object for the document body.

          let range = context.document.getSelection();
          range.load()
          context.trackedObjects.add(range);
          // Queue a commmand to get the OOXML contents of the body.
          // var bodyOOXML = body.getOoxml();

          // Synchronize the document state by executing the queued commands,
          // and return a promise to indicate task completion.
          return context.sync(range)
      }).then(function (range) {
          range.insertOoxml(
              '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
              '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512" >' +
              '<pkg:xmlData ><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" >' +
              '<Relationship Id="rId1" Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" /></Relationships>' +
              '</pkg:xmlData ></pkg:part><pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256">' +
              '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />' +
              '</Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
              '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
              '<w:body><w:p><w:commentRangeStart w:id="0"/><w:r>' +
              '          <w:sdt>' +
              '              <w:sdtPr>' +
              '                <w:alias w:val="MyContentControlTitle"/>' +
              '                <w:id w:val="1382295294"/>' +
              '                <w15:appearance w15:val="hidden"/>' +
              '              </w:sdtPr>' +
              '              <w:sdtContent>' +
              '                  <w:r>' +
              '                  <w:t>' + range.text + '</w:t>' +
              '                </w:r>' +
              '              </w:sdtContent>' +
              '            </w:sdt>' +
              '</w:r><w:commentRangeEnd w:id="0"/>' +
              '<w:commentReference w:id="0"/></w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
              '<pkg:part xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">' +
              '<pkg:xmlData><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
              '<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="0">' +
              '<w:p>' +
              '<w:r>' +
              '<w:t>Comment</w:t>' +
              '</w:r>' +
              '<w:hyperlink Id="rId2" w:history="1">' +
              '<w:r>' +
              '<w:rPr>' +
              '<w:rStyle w:val="Hyperlink"/>' +
              '</w:rPr>' +
              '<w:t>Google</w:t>' +
              '</w:r>' +
              '</w:hyperlink>' +
              '</w:p>' +
              '</w:comment></w:comments>' +
              '</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/comments.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
              '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships></pkg:xmlData></pkg:part></pkg:package>',
              "Replace"
          );

          // myCC.appearance = "Tags"; // you can also set it to 'boundingBox' or  'tags'
          // myCC.title = "BindingID"
          // console.log(myCC)

           //Office.context.addHandlerAsync(myCC.onDataChanged, changeMe())
           //document.addHandlerAsync(myCC.onSelectionChanged, touchMe())


          function touchMe() {
              console.log("Don't touch me!")
          }

          function ChangeMe(eventData) {
              console.log("Eeew")
              console.log(eventData)
          }
          return range.context.sync();
      }).then(function (range) {
          Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", {}, function (result) {
              console.log(result.status);
              if (result.status == "succeeded") {
                  // lets create an event!
                  result.value.addHandlerAsync(Office.EventType.BindingSelectionChanged, function () {
                      console.log("Are you looking at me?");
                  })
                  result.value.addHandlerAsync(Office.EventType.BindingDataChanged, function (evt) {
                      console.log("Don't touch me you pervert!");
                      console.log(evt)
                  })
              }
          });
          range.context.trackedObjects.remove(range);
          return context.sync();
      })
  }