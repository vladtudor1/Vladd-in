/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import Vue from 'vue/dist/vue.esm.js';

Vue.config.productionTip = false;

const app = new Vue({
  el: '#app',
  data: {
    document: {},
    ooxml: 'Document OOXML',
    bookmarkList: [
      {
        id: 0,
        name: "Voorwoord",
        isSelected: false
      },
      {
        id: 1,
        name: "Samenvatting",
        isSelected: false
      },
      {
        id: 2,
        name: "Inleiding",
        isSelected: false
      },
      {
        id: 3,
        name: "Begroting",
        isSelected: false
      }
    ]
  },
  computed: {
    ooxmlBody: function() {
      return this.ooxml.substring(this.ooxml.indexOf('<w:body>'), this.ooxml.indexOf('</w:body>'))
    }
  },
  mounted: function () {
    var _this = this;
    Office.initialize = function (reason) {
      OfficeExtension.config.extendedErrorLogging = true;
      _this.document = Office.context.document;
      _this.tryUpdatingSelectedWord();
      _this.document.addHandlerAsync("documentSelectionChanged", _this.tryUpdatingSelectedWord);
      _this.updateOOXML();
    }
  },
  methods: {
    setCheck: function(bookmarkItem){
      bookmarkItem.isSelected = true
    },
    unsetCheck: function(bookmarkItem){
      bookmarkItem.isSelected = false
    },
    tryUpdatingSelectedWord: function () {
      this.document.getSelectedDataAsync(Office.CoercionType.Text, this.selectedTextCallback);
    },
    selectedTextCallback: function (selectedText) {
      this.selectedText = $.trim(selectedText.value);
    },
    updateOOXML: function() {
      var _this = this;
      Word.run(async function (ctx) {
        var docOoxml = ctx.document.body.getOoxml();
        return ctx.sync().then( function() {
          _this.ooxml = docOoxml.value
        });
      })
    },
    insertBookmark: function (bookmarkItem) {
      var _this = this;
      console.log('init')
      console.log('bookmarking '+bookmarkItem.name+' lol')
      Word.run(async function (context) {
        console.log(context)
        // Create a bookmark ID - 0 works just fine as Word will generate a new number upon insert which is awesome!
        var bkmkId = 0;

        // Create a bookmark Name - it must be unique or it will overwrite a current one!
        // It must be no longer than 40 characters!
        // Please see the behavior here:
        var bkmkName = "_TOC_MANUAL_" + bookmarkItem.name;

        // Queue a command to get the current selection and then
        // create a proxy range object with the results.
        var range = context.document.getSelection();

        // Use the extension to insert the bookmark.  It really should be this easy.  See the header of the
        // extension for an idea of the current suggested API for Microsoft.  I only wrote a quick insertBookmark
        // method that doesn't quite fit the signature for what's proposed btw.  I wanted this to be quick and
        // minimal.  Plus, the extension should really extend range, document.body, etc.
        _this.insertOOXMLBookmark(bkmkId, bkmkName, _this.handleSuccess, _this.errorHandler);
      });
      this.updateOOXML()
    },
    findBookmark: function(bookmarkItem) {
      Word.run(async function (context) {
        console.log(context.document.getBookmarkRange('_TOC_MANUAL_'+bookmarkItem.name).getOoxml())
        await context.sync();
      });
    },
    deleteBookmark: function (bookmarkItem) {
      var W = openXml.W;
      var XElement = Ltxml.XElement;
      Word.run(async function(context){
        //Define the range and OOXML of the selection 
        console.log("Deleting bookmark " + bookmarkItem.name)
        context.document.deleteBookmark('_TOC_MANUAL_'+bookmarkItem.name)
        return context.sync();
      });
      this.updateOOXML()
    },
    insertOOXMLBookmark:function (bkmkId, bkmkName, handleSuccess, handleError) {
      var _this = this;
      // OpenXml...
      var XAttribute = Ltxml.XAttribute;
      var XElement = Ltxml.XElement;
      //var XDocument = Ltxml.XDocument;
      var W = openXml.W;

      Word.run(function (ctx) {
          // Queue a command to get the current selection and then create a proxy range object with the results.
          var range = ctx.document.getSelection().paragraphs.getFirstOrNullObject()
          // sync with our document to conflate the ooxml...
          var ooxml = range.getOoxml()
          return ctx.sync().then(function () {
              // now we can use the ooxml var
              // open the full OPC package we get in the openXml SDK
              var doc = new openXml.OpenXmlPackage(ooxml.value);

              // Create the bookmarkStart/End elements.
              // <w:bookmarkStart w:id="{#}" w:name="{label}"/>
              // <w:bookmarkEnd w:id="{#}"/>

              var bookmarkStart = new XElement("w:bookmarkStart",
                  new XAttribute("w:id", bkmkId),
                  new XAttribute("w:name", bkmkName));
              var bookmarkEnd = new XElement("w:bookmarkEnd",
                  new XAttribute("w:id", bkmkId));

              // Parse the OPC which represents the selection as an entire document
              var mainPart = doc.mainDocumentPart();
              var mainPartXDoc = mainPart.getXDocument();

              // CAVEAT: Need to confirm Word's behavior! Bookmarks need to be sub-elements of paragraphs.  It
              // seems that Word wraps selection content in paragraph markup when necessary so that we always
              // get a valid block.  In other words, it seems that getOoxml was designed with copy/paste like functionality
              // in mind in that you always get valid Opc (despite the name of the method).

              // TODO: research if I should use .elements() instead of .descendantNodes()?
              // Find first paragraph and add the bookmarkStart - the selection is wrapped in a paragraph
              var body = mainPartXDoc.root.element(W.body);
              var functorPMatch = function (e) {
                  // locates paragraph markup
                  return (e != null) && !(e.name === undefined) && (e.name != null) && (e.name.localName === "p");
              };

              var nodeFirstP = body.descendantNodes().first(functorPMatch);
              if (nodeFirstP === undefined) {
                  // TODO - No paragraphs - we should never get here.
                  console.log("Word-extensions.js: This should never happen; we don't have a paragraph in the selection.")
              }
              else {
                  // Insert the bookmark start here.
                  nodeFirstP.addFirst(bookmarkStart);


                  // Now locate the last paragraph, as there may be multiples, and insert our end after the end of it.
                  var nodeLastP = body.descendantNodes().last(functorPMatch);
                  if (nodeLastP === undefined) {
                      // No paragraphs
                  }
                  else {
                      // Now locate the last paragraph and insert our end after the end of it.
                      var count = nodeLastP.descendantNodes().count();
                      if (0 == count) {
                          nodeLastP.addFirst(bookmarkEnd);
                      } else {
                          var node = nodeLastP.getLastNode();

                          // WORD BUG - Anchor Bookmarks - 
                          //      <w:bookmarkStart w:id="0" w:name="uniquename"/><w:bookmarkEnd w:id="0"/>
                          // Word has a bug that does not allow anchor bookmarks to be inserted.  Rather than
                          // recognize them on the insertOoxml invocation below, it simply throws them away.  Ugh!

                          // WORD BUG WORKAROUND - insert some text seems to be the only way to fix it in Word.
                          // Unfortunately, that causes another paragraph to be entered and users need to delete
                          // the extra character and paragraph and the bookmark will still be retained.  Word
                          // seemingly has special logic when it creates these anchor type of bookmarks from it's
                          // interface, because it does work there.
                          if (1 == count) {
                              // We're an empty anchor bookmark.  There's a bug in Word it seems where it won't insert the
                              // bookmark without some node in between so we'll create a work around here.
                              //var gbid = bkmkId + 1;
                              //var goBackBookmarkkStart = new XElement("w:bookmarkStart",
                              //                        new XAttribute("w:id", gbid),
                              //                        new XAttribute("w:name", "_GoBack"));
                              //var goBackBookmarkEnd = new XElement("w:bookmarkEnd",
                              //    new XAttribute("w:id", gbid));

                              // Putting in some non-whitespace text effectively fixes the bug; We'll add a space...
                              // ...interestingly, Word will add an extra paragraph; which is another BUG in my
                              // book!
                              var e = new XElement("w:r", new XElement("w:t", '[delete]'));
                              node.addAfterSelf(e);

                              //node.addAfterSelf(goBackBookmarkEnd);
                              node.addAfterSelf(bookmarkEnd);
                              //node.addAfterSelf(goBackBookmarkkStart);
                          }
                          else {
                              // Normal case
                              node.addAfterSelf(bookmarkEnd);
                          }
                      }
                  }
              }

              // Now queue up a reinsert and replace the current selection.  We've been careful above to only
              // insert our new bookmark items and not tamper with any other OOXML so all other structures should
              // be retained.
              range.insertOoxml(doc.saveToFlatOpc(), Word.InsertLocation.replace);

              return ctx.sync()
          // Synchronize the document state by executing the queued commands.
          .then(ctx.sync)
          .then(function () {
            handleSuccess();
          })
          .catch(function (error) {
            handleError(error);
          });
        });
      });
    },
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    errorHandler: function (error) {
      // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
      this.showNotification("Error:", error);
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    },
    handleSuccess: function() {
    },
    // Helper function for displaying notifications
    showNotification: function (header, content) {

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
    },
  }
})
