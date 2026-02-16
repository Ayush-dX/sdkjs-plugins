(function(window, undefined){

    // 1. Initialize
    window.Asc.plugin.init = function() {
        // Event: Fired when cursor moves. We check if text is selected.
        window.Asc.plugin.attachEvent("onTargetPositionChanged", function(){
            window.Asc.plugin.executeMethod("GetSelectedText", [], function(text) {
                var hasSelection = (text && text.length > 0);
                
                // Communicate status back to the UI (index.html)
                // We target the iframe by the GUID defined in config.json
                var iframe = document.getElementById("iframe_asc.{B8D8C300-1234-5678-9ABC-DEF123456789}");
                if(iframe){
                    iframe.contentWindow.postMessage({
                        type: "SELECTION_STATUS",
                        hasSelection: hasSelection,
                        text: text
                    }, "*");
                }
            });
        });
    };

    // 2. Required for plugins (even if unused in Panel mode)
    window.Asc.plugin.button = function(id) {
        this.executeCommand("close", "");
    };

    // 3. Custom: Find Header by Name and Highlight it
    window.Asc.plugin.findAndHighlightHeader = function(headerName) {
        this.callCommand(function() {
            var oDocument = Api.GetDocument();
            var aParagraphs = oDocument.GetAllParagraphs();
            var rangeStart = null, rangeEnd = null;
            var found = false;
            var headerText = "";

            // Loop paragraphs to find Header text
            for (var i = 0; i < aParagraphs.length; i++) {
                var oPara = aParagraphs[i];
                var sText = oPara.GetText();
                
                if (!found && sText.toLowerCase().includes(headerName.toLowerCase())) {
                    rangeStart = oPara.GetRange().GetStartPos();
                    found = true;
                    headerText = sText;
                } else if (found && oPara.GetStyle().GetName().indexOf("Heading") === 0) {
                    // Stop at the next heading
                    rangeEnd = aParagraphs[i - 1].GetRange().GetEndPos();
                    break;
                }
            }
            if (found && !rangeEnd) {
                rangeEnd = aParagraphs[aParagraphs.length - 1].GetRange().GetEndPos();
            }

            if (found) {
                var oRange = oDocument.GetRange(rangeStart, rangeEnd);
                oRange.SetHighlight("yellow");
                return {
                    found: true,
                    text: oRange.GetText(),
                    headerTitle: headerText,
                    start: rangeStart,
                    end: rangeEnd
                };
            }
            return { found: false };

        }, true, true, function(result){
            var iframe = document.getElementById("iframe_asc.{B8D8C300-1234-5678-9ABC-DEF123456789}");
            if(iframe) iframe.contentWindow.postMessage({ type: "HEADER_CAPTURED", payload: result }, "*");
        });
    };

    // 4. Custom: Highlight the currently selected text
    window.Asc.plugin.highlightSelection = function() {
        this.callCommand(function() {
            var oDocument = Api.GetDocument();
            var oRange = oDocument.GetRangeBySelect();
            oRange.SetHighlight("cyan");
            return {
                text: oRange.GetText(),
                start: oRange.GetStartPos(),
                end: oRange.GetEndPos()
            };
        }, true, true, function(result){
            var iframe = document.getElementById("iframe_asc.{B8D8C300-1234-5678-9ABC-DEF123456789}");
            if(iframe) iframe.contentWindow.postMessage({ type: "SELECTION_CAPTURED", payload: result }, "*");
        });
    };

    // 5. Custom: Replace text (Backend Action)
    window.Asc.plugin.replaceContextText = function(start, end, newText) {
        this.callCommand(function() {
            var oDocument = Api.GetDocument();
            var oRange = oDocument.GetRange(start, end);
            oRange.SetHighlight("None"); // Remove highlight
            oRange.SetText(newText);     // Replace
        }, true);
    };

    // 6. Custom: Remove Highlight (Cleanup)
    window.Asc.plugin.clearHighlight = function(start, end) {
        this.callCommand(function() {
            var oDocument = Api.GetDocument();
            var oRange = oDocument.GetRange(start, end);
            oRange.SetHighlight("None");
        }, true);
    };

})(window, undefined);