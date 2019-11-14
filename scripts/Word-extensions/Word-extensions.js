// ==================================================================================
// Word-extensions.js
//
// Copyright (c) 2016 ronjonesjr.  Licensed under the MIT license.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// Mimic extending the Word.js object with functionality that is missing. 
//
// BUG ALERT: There's a bug in Word 2016 (16.0.6925.1018) noted below.  Earlier versions
// of Word 2016 have a few other glitches in them, so if you're running an earlier 
// version than mine, you may also see an additional paragraph inserted.  That's
// even worse behavior, but the team at Microsoft must have fixed that already!
//
// Here's what I would like to see Microsoft add...
//
// Bookmark: is missing altogether
// --------------------------------
// ++ text
// ++ id
// ++ name
// ++ bookmarks - type:BookmarkCollection - gets the sub-bookmarks
// ++ inlinePictures
// ++ paragraphs
// ++ parentBookmark
// ++ select
// ++ type - public or private
// ++ ... see other methods, properties, etc. on ContentControl (https://dev.office.com/reference/add-ins/word/contentcontrol)
//
// BookmarkCollection: is missing altogether
// ------------------------------------------
// ++ items
// ++ getById(id:number)
// ++ getByName(name:string)
// ++ load(param:object)
// ++ ...see other methods, properties, etc. on ContentControlCollection (https://dev.office.com/reference/add-ins/word/contentcontrolcollection)
//
// Range: should have additional methods
// ---------------------------------------
// ++ insertBookmark([id,] name)
// ++ BookmarkCollection - read only
// ++ parentBookmark - read only
//
// ==================================================================================

// I'll only mimic what I can and what I need below for now.  Note, there won't be a
// programmatic way to "select" a bookmark in the document since we can't create a
// range...which is by design from Microsoft.  Word can do that from the
// bookmark menu but I won't be able to do that from an add-In unfortunately.  I can
// insert them though...the code below leverages the OpenXml.js code from 
// http://ericwhite.com to do exactly that.  This is a poor workaround for the missing
// functionality in Office.js though..thus there are a few caveats below and a few
// bugs in Microsoft Word 2016....all are called out below.
//
// References
// http://ericwhite.com/blog/blog/open-xml-sdk-for-javascript/
// http://openxmlsdkjs.codeplex.com/
// http://ericwhite.com/blog/forums/forum/open-xml-sdk-javascript
//
//<!-- This is the openxmlsdk from http://ericwhite.com -->
//<script src="../../Scripts/openxmlsdkjs-01-01-02/linq.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/ltxml.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/ltxml-extensions.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/jszip.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/jszip-load.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/jszip-inflate.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/jszip-deflate.js"></script>
//<script src="../../Scripts/openxmlsdkjs-01-01-02/openxml.js"></script>
//

var officeWordExtension = (function () {
    "use strict";

    var officeWordExtension = {
    };

    // insert a bookmark at the current selection
    officeWordExtension.insertBookmark = function (bkmkId, bkmkName, handleSuccess, handleError) {
        // OpenXml...
        var XAttribute = Ltxml.XAttribute;
        var XElement = Ltxml.XElement;
        //var XDocument = Ltxml.XDocument;
        var W = openXml.W;

        Word.run(function (ctx) {
            // Queue a command to get the current selection and then create a proxy range object with the results.
            var range = ctx.document.getSelection();

            // Get our OOXML...
            var ooxml = range.getOoxml();

            // sync with our document to conflate the ooxml...
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

                return ctx.sync().then(function () {
                    console.log('Word-extensions.js: Bookmark inserted {name: "' + bkmkName + '" id: "' + bkmkId + '"}');
                });
            })
            // Synchronize the document state by executing the queued commands.
            .then(ctx.sync)
            .then(function () {
                handleSuccess();
            })
            .catch(function (error) {
                handleError(error);
            });
        });
    };

    return officeWordExtension;
})();
