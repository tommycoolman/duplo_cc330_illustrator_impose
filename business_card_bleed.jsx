var docRef = app.activeDocument;

// Total art size should be 3.75x2.2"

var i = docRef.artboards.getActiveArtboardIndex();
var rect = app.activeDocument.pathItems.rectangle (docRef.artboards[i].artboardRect[1],docRef.artboards[i].artboardRect[0], 252, 144);

rect.fillColor = rect.strokeColor = new NoColor();
rect.selected = true;

app.executeMenuCommand("unlockAll");
app.executeMenuCommand("selectallinartboard");
app.executeMenuCommand("group");

var doc = app.activeDocument.selection;

var docPreset = new DocumentPreset;
    docPreset.colorMode = DocumentColorSpace.CMYK;
    docPreset.title  = "New File 01";
    docPreset.width  = 864;
    docPreset.height = 1296;

var presetArt = app.startupPresetsList[3];

var newdoc = app.documents.addDocument(presetArt, docPreset);

if(doc.length > 0){
    if(doc.length > 0){
        var card = new Array();
        for(var x=0;x<doc.length;x++){
            doc[x].selected = false;

            var cardCount = 0;
            var topOffset = 1296-25.2+7.2;
            var leftOffset = 36-7.2;
            for(var col=0;col<3;col++){
                for(var row=0;row<8;row++){
                    card[cardCount] = doc[x].duplicate(newdoc,ElementPlacement.PLACEATEND);
                    card[cardCount].left = leftOffset;
                    card[cardCount].top = topOffset;
                    topOffset -= (2.2 * 72);
                    cardCount++;
                }
                leftOffset += (3.75*72);
                topOffset = 1296-25.2+7.2;
            }

        }
    }
    else{
        doc.selected = false;
        newitem = doc.parent.duplicate(newdoc,ElementPlacement.PLACEATEND);
    }
}
