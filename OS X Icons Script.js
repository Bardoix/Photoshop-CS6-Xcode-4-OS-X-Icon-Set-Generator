#target photoshop

function main() {
    if (app.documents.length === 0) {
        alert("Nessun documento aperto! Apri l'icona sorgente prima di eseguire lo script.");
        return;
    }

    var doc = app.activeDocument;

    try {
        var path = doc.path;
    } catch (e) {
        alert("Salva il documento prima di eseguire lo script per definire la cartella di destinazione delle icone.");
        return;
    }

    var originalDisplayDialogs = app.displayDialogs;
    app.displayDialogs = DialogModes.NO;

    var iconSizes = [
        ["icon_16x16.png", 16],
        ["icon_16x16@2x.png", 32],
        ["icon_32x32.png", 32],
        ["icon_32x32@2x.png", 64],
        ["icon_128x128.png", 128],
        ["icon_128x128@2x.png", 256],
        ["icon_256x256.png", 256],
        ["icon_256x256@2x.png", 512],
        ["icon_512x512.png", 512],
        ["icon_512x512@2x.png", 1024]
    ];

    var savedState = doc.activeHistoryState;

    var pngOptions = new PNGSaveOptions();
    pngOptions.interlaced = false;

    for (var i = 0; i < iconSizes.length; i++) {
        var name = iconSizes[i][0];
        var size = iconSizes[i][1];
        var destFile = new File(path + "/" + name);
        doc.resizeImage(UnitValue(size, "px"), UnitValue(size, "px"), null, ResampleMethod.BICUBICSHARPER);
        doc.saveAs(destFile, pngOptions, true, Extension.LOWERCASE);
        doc.activeHistoryState = savedState;
    }
    app.displayDialogs = originalDisplayDialogs;
    alert("Esportazione completata!\nLe 10 icone PNG sono state salvate in:\n" + path);
}

main();