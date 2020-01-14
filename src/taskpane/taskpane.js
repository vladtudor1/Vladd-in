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
    xmlObj: {},
    bookmarkList: [
      {
        id: 0,
        doc_order: 1,
        position: null,
        name: "Voorwoord",
        isSelected: false,
        outOfOrder:false
      },
      {
        id: 1,
        doc_order: 2,
        position: null,
        name: "Samenvatting",
        isSelected: false,
        outOfOrder:false
      },
      {
        id: 2,
        doc_order: 3,
        position: null,
        name: "Inleiding",
        isSelected: false,
        outOfOrder:false
      },
      {
        id: 3,
        doc_order: 4,
        position: null,
        name: "Begroting",
        isSelected: false,
        outOfOrder:false
      }
    ]
  },
  computed: {
    ooxmlBody: function() {
      return this.ooxml.substring(this.ooxml.indexOf('<w:document'), this.ooxml.indexOf('</w:document'))
      // return this.ooxml.substring(this.ooxml.indexOf('<w:'+element), this.ooxml.indexOf('</w:'+element))
    },
  },
  mounted: function () {
    var _this = this;
    Office.initialize = function (reason) {
      OfficeExtension.config.extendedErrorLogging = true;
      _this.document = Office.context.document;
      _this.syncOOXML();
      _this.document.addHandlerAsync("documentSelectionChanged", _this.syncOOXML);
    }
  },
  updated: function() {
    var _this = this;
    var maxPos = 0
    var parser = new DOMParser();
    var xmlDoc = parser.parseFromString(_this.ooxml,"text/xml");

    $.each(this.bookmarkList, function(key, value)  {
        // try to find bookmark
        var position = _this.ooxml.indexOf('_TOC_MANUAL_'+value.name) 
        // todo 1: zoeken met een Xpath in de XMLdoc
        // //bookmarkStart[@w:name='_TOC_MANUAL_'+value.name]
        _this.bookmarkList[key].position = position;
        // todo 2: sla in het bookmarkList VueJS object het absolute Xpath pad op als xpathObj['start']
        //      /w:document[1]/w:body[1]/w:p[xx]/w:r[xx]
        if (position !== -1) {
          //bookmark is found
          _this.bookmarkList[key].isSelected = true;
        }
        // if bookmark is not in OOXML:
        else {
          _this.bookmarkList[key].isSelected = false;
        }
        if (value.position < maxPos) {
          _this.bookmarkList[key].outOfOrder = true;
        }
        else {
          _this.bookmarkList[key].outOfOrder = false;
          maxPos = Math.max(maxPos, _this.bookmarkList[key].position)
        }
    });

    console.log(xmlDoc)
  },
  methods: {
    dumpState: function() {
      console.log(this._data.bookmarkList)
    },
    syncOOXML: function () {
      var _this = this;
      Word.run(async function (ctx) {
        var docOoxml = ctx.document.body.getOoxml();
        return ctx.sync().then( function() {
          _this.ooxml = docOoxml.value
        });
      })
    },
    highlightSelection: function (bookmarkItem){

      console.log(bookmarkItem.name)

      Word.run( function(context){
        
        var documentOoxml = context.document.body.getOoxml();

        return context.sync().then(function(){

        var ooxml = documentOoxml.value;
        var bookmarkName = '_TOC_MANUAL_'+bookmarkItem.name
        var present = ooxml.indexOf('w:name="'+bookmarkName)

          if (present !== -1){

            var bookmarkObject = context.document.getBookmarkRangeOrNullObject(bookmarkName)
            var bookmarkRange = bookmarkObject.load();

            // Then add the highlight to the selected bookmark range
            bookmarkRange.select('select')
            return context.sync().then(function(){
            })
          } else {
            console.log("Nothing was found")
            var range = context.document.getSelection();
            range.select('start');
            return context.sync();
          }
        })
      })

    },
    insertBookmark: function (bookmarkItem) {
      var _this = this;
      var bookmarkList = _this.bookmarkList;
      console.log('init')
      console.log('bookmarking '+ bookmarkItem.name +' lol')
      bookmarkItem.isSelected = true;
      console.log(bookmarkList)
      Word.run(async function (context) {
        // console.log(context)
        // Create a bookmark ID - 0 works just fine as Word will generate a new number upon insert which is awesome!
        var bkmkId = 0;

        // Create a bookmark Name - it must be unique or it will overwrite a current one!
        // It must be no longer than 40 characters!
        // Please see the behavior here:
        var bkmkName = "_TOC_MANUAL_" + bookmarkItem.name;

        // Queue a command to get the current selection and then
        // create a proxy range object with the results.
        var range = context.document.getSelection();

        _this.insertOOXMLBookmark(bkmkId, bkmkName, _this.handleSuccess, _this.errorHandler);
      });
      this.syncOOXML()
    },
    
    deleteBookmark: function (index) {
      var bookmarkItem = this.bookmarkList[index];

      Word.run(async function(context){
        console.log("Deleting bookmark " + bookmarkItem.name)
        context.document.deleteBookmark('_TOC_MANUAL_'+ bookmarkItem.name)
        return context.sync();
      });
      this.syncOOXML()
    },
    insertOOXMLBookmark:function (bkmkId, bkmkName, handleSuccess, handleError) {
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
