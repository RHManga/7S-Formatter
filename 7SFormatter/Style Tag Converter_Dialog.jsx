// Main function
function main() {
    // Get a reference to the document
    var doc = app.activeDocument;

    // Get all layers and character styles
    var layers = doc.layers.everyItem().name;
    var characterStyles = doc.characterStyles.everyItem().name;

    // Define default selections for each dropdown menu
    var defaultLayerIndex = 1; // Change this to the index of the default layer
    var defaultBoldStyleIndex = 4; // Change this to the index of the default bold style
    var defaultItalicStyleIndex = 3; // Change this to the index of the default italic style
    var defaultBoldItalicStyleIndex = 5; // Change this to the index of the default bold italic style

    // Create a dialog box
    var dialog = new Window("dialog", "Apply Character Styles");

    // Add labels for each dropdown menu
    var textLayerLabel = dialog.add("statictext", undefined, "Text Layer:");
    var layerDropdown = dialog.add("dropdownlist", undefined, layers);
    layerDropdown.selection = defaultLayerIndex; // Set default selection

    var boldLabel = dialog.add("statictext", undefined, "Bold Style:");
    var boldDropdown = dialog.add("dropdownlist", undefined, characterStyles);
    boldDropdown.selection = defaultBoldStyleIndex; // Set default selection

    var italicLabel = dialog.add("statictext", undefined, "Italic Style:");
    var italicDropdown = dialog.add("dropdownlist", undefined, characterStyles);
    italicDropdown.selection = defaultItalicStyleIndex; // Set default selection

    var boldItalicLabel = dialog.add("statictext", undefined, "Bold Italic Style:");
    var boldItalicDropdown = dialog.add("dropdownlist", undefined, characterStyles);
    boldItalicDropdown.selection = defaultBoldItalicStyleIndex; // Set default selection

    // Add buttons
    dialog.add("button", undefined, "OK");
    dialog.add("button", undefined, "Cancel");

    // Show the dialog box
    if (dialog.show() === 1) { // User clicked OK
        // Get the selected layer and character styles
        var selectedLayer = layers[layerDropdown.selection.index];
        var selectedBoldStyle = characterStyles[boldDropdown.selection.index];
        var selectedItalicStyle = characterStyles[italicDropdown.selection.index];
        var selectedBoldItalicStyle = characterStyles[boldItalicDropdown.selection.index];

        // Find the target layer
        var targetLayer = doc.layers.itemByName(selectedLayer);

        // Check if the layer exists
        if (targetLayer.isValid) {
            // Define the doScript function to be executed
            var doScriptFunction = function() {
                // Iterate through all text frames in the layer
                for (var i = 0; i < targetLayer.textFrames.length; i++) {
                    var textFrame = targetLayer.textFrames[i];
                    var text = textFrame.contents;

                    // Use regular expressions to find pairs of tags '[b][/b]', '[i][/i]', and '[k][/k]'
                    var boldRegex = /\[b\]([^\]]*?)\[\/b\]/g;
                    var italicRegex = /\[i\]([^\]]*?)\[\/i\]/g;
                    var boldItalicRegex = /\[k\]([^\]]*?)\[\/k\]/g;
                    var boldMatch, italicMatch, boldItalicMatch;

                    // Apply bold character style
                    while (boldMatch = boldRegex.exec(text)) {
                        var startIndex = boldMatch.index + 3; // Length of '[b]'
                        var endIndex = boldMatch.index + boldMatch[1].length + 2; // Length of the captured text

                        textFrame.characters.itemByRange(startIndex, endIndex).appliedCharacterStyle = doc.characterStyles.itemByName(selectedBoldStyle);
                    }

                    // Apply italic character style
                    while (italicMatch = italicRegex.exec(text)) {
                        var startIndex = italicMatch.index + 3; // Length of '[i]'
                        var endIndex = italicMatch.index + italicMatch[1].length + 2; // Length of the captured text

                        textFrame.characters.itemByRange(startIndex, endIndex).appliedCharacterStyle = doc.characterStyles.itemByName(selectedItalicStyle);
                    }

                    // Apply bold italic character style
                    while (boldItalicMatch = boldItalicRegex.exec(text)) {
                        var startIndex = boldItalicMatch.index + 3; // Length of '[k]'
                        var endIndex = boldItalicMatch.index + boldItalicMatch[1].length + 2; // Length of the captured text

                        textFrame.characters.itemByRange(startIndex, endIndex).appliedCharacterStyle = doc.characterStyles.itemByName(selectedBoldItalicStyle);
                    }
                }
            };

            // Execute the doScript function wrapped in an undoable action
            app.doScript(doScriptFunction, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Apply Character Styles");
        } else {
            alert("Layer '" + selectedLayer + "' not found.");
        }
    }
}

// Call the main function
main();
