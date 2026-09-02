var doc = app.activeDocument;

var carpeta = Folder.selectDialog("Selecciona la carpeta on hi ha les fotos");
var archivoWord = File.openDialog("Selecciona el document de Word amb els peus de foto", "Archivos de Word:*.docx;*.doc;");

if (!carpeta || !archivoWord) {
    alert("Operació cancel·lada: Cal seleccionar una carpeta i un fitxer.");
    exit();
}

var win = new Window("dialog", "Paràmetres d'entrada");
win.alignChildren = ["fill", "top"];
win.spacing = 8;

function addRow(parent, labelText) {
    var grp = parent.add("group");
    grp.orientation = "row";
    grp.alignChildren = ["left", "center"];
    grp.spacing = 10;
    if (labelText) grp.add("statictext", undefined, labelText);
    return grp;
}

var rowFormat = addRow(win, "format: ");
var inputStr1 = rowFormat.add("edittext", undefined, "Fig ");
inputStr1.preferredSize.width = 60;
rowFormat.add("statictext", undefined, "1");
var inputStr2 = rowFormat.add("edittext", undefined, "");
inputStr2.preferredSize.width = 60;

var rowFirst = addRow(win, "num. primera figura");
var inputNum1 = rowFirst.add("edittext", undefined, "1");
inputNum1.preferredSize.width = 60;

var rowLast = addRow(win, "num. última figura");
var inputNum2 = rowLast.add("edittext", undefined, "10");
inputNum2.preferredSize.width = 60;

var rowStyle = addRow(win, "Nom de l'estil de paràgraf");
var inputStr3 = rowStyle.add("edittext", undefined, "6. peu foto");
inputStr3.preferredSize.width = 120;

var rowWidth = addRow(win, "Ample predeterminat: ");
var inputNum3 = rowWidth.add("edittext", undefined, "80");
inputNum3.preferredSize.width = 60;

var rowCheck = addRow(win, "Preguntar a cada figura?");
var chkBoxOption = rowCheck.add("checkbox", undefined, "");
chkBoxOption.value = false;

var btnGroup = win.add("group");
btnGroup.alignment = ["right", "center"];
btnGroup.add("button", undefined, "D'acord", {name: "ok"});
btnGroup.add("button", undefined, "Cancel·lar", {name: "cancel"});

if (win.show() !== 1) {
    exit(); 
}

var pre = inputStr1.text;
var post = inputStr2.text;
var nombreEstilo = inputStr3.text;

var primeraCrida = parseInt(inputNum1.text, 10) || 0;
var ultimaCrida = parseInt(inputNum2.text, 10) || 0;
var rawText = inputNum3.text.replace(",", ".");
var anchoDeseado = parseFloat(rawText) || 0;
var NOusarAmplePredeterminat = chkBoxOption.value; // Returns true or false
var anchoDeseadoOriginal = anchoDeseado;

var myExtensions = [".png", ".jpg", ".jpeg", ".tif", ".tiff", ".gif", ".pdf", ".psd", ".ai", ".heic", ".jfif"];

function extraerNumero(texto) {
    if (!texto) return null;
    
    var str = String(texto) + "";
    var digitosEncontrados = "";
    var leyendoNumero = false;

    for (var i = 0; i < str.length; i++) {
        var cchar = str[i];
        
        if (cchar >= '0' && cchar <= '9') {
            digitosEncontrados += cchar;
            leyendoNumero = true;
        } else if (leyendoNumero) {
            break;
        }
    }
    return digitosEncontrados !== "" ? parseInt(digitosEncontrados, 10) : null;
}

function obtenerImagenesMultiples(myFolder) {
    if (File.fs == "Macintosh") {
        return myFolder.getFiles(myFileFilterMac);
    } else {
        var myFilteredFiles = new Array();
        for (var i = 0; i < myExtensions.length; i++) {
            var myFiles = myFolder.getFiles("*" + myExtensions[i]);
            if (myFiles.length != 0) {
                for (var j = 0; j < myFiles.length; j++) {
                    myFilteredFiles.push(myFiles[j]);
                }
            }
        }
        return myFilteredFiles;
    }
}

function myFileFilterMac(myFile) {
    var myFileType = myFile.type;
    switch (myFileType) {
        case "JPEG":
        case "PNGf":
        case "TIFF":
            return true;
        default:
            var fileName = myFile.name.toLowerCase();
            for (var i = 0; i < myExtensions.length; i++) {
                if (fileName.indexOf(myExtensions[i]) > -1) {
                    return true;
                }
            }
    }
    return false;
}


if (carpeta != null) {
    var fotos = obtenerImagenesMultiples(carpeta);

    if (fotos.length > 0) {
        
        var tempFrame = doc.pages[0].textFrames.add();
        tempFrame.geometricBounds = [0, 0, 200, 200];
        tempFrame.place(archivoWord);
        
        for (var numCrida = primeraCrida; numCrida <= ultimaCrida; numCrida++) {
            var estiloPie = doc.paragraphStyles.itemByName(nombreEstilo);
            var textBuscar = pre + numCrida + post;
            var top = 40;
            var left = 40;

            doc.recompose();
            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.findWhat = textBuscar;
            var resultados = doc.findText();
            if (resultados.length === 0) {
                alert("no s'ha trobat la crida " + textBuscar)
                continue;
            }

            var resultadoEncontrado = resultados[0];
            var puntoDeInsercion = resultadoEncontrado.insertionPoints[-1];
            var parentFrames = puntoDeInsercion.parentTextFrames;
            
            if (!parentFrames || parentFrames.length === 0) {
                alert("Atenció: La crida " + textBuscar + " està en text desbordat o no té un marc assignat.");
                continue; 
            }

            var textFramePadre = parentFrames[0];
            var page = textFramePadre.parentPage;

            var flag1 = false;
            var j = 0;
            while (j < fotos.length && !flag1){

                var foto = fotos[j];
                var nomFoto = String(decodeURIComponent(foto.name));
                var numFoto = extraerNumero(nomFoto)

                if (numFoto == numCrida){
                    flag1 = true;
                } else {
                    j++;
                }
            }
            if (!flag1) {
                alert("foto no trobada " + numCrida)
                continue;
            }

            var flag2 = false;
            var parrafosTemp = tempFrame.paragraphs;
            var k = 0;
            var parrafoEncontrado = null;

            while (k < parrafosTemp.length && !flag2) {
                var rawText = parrafosTemp[k].contents;
                var txtParrafo = rawText.replace(/[\r\n\t]+$/, "").replace(/^\s+|\s+$/g, "");
                var numPeu = extraerNumero(txtParrafo);

                if (numPeu == numCrida) {
                    flag2 = true;
                    parrafoEncontrado = parrafosTemp[k];
                } else {
                    k++;
                }
            }
            if (!flag2) {
                alert("peu no trobat " + numCrida)
            }

            var colocados = page.place(foto, [left, top]);
            var imagen = colocados[0];
            var marco = imagen.parent;

            // var anchoDeseado = 80; // en pt ample de les fotos
            if (NOusarAmplePredeterminat){
                var win1 = new Window("dialog", "Figura " + numCrida);
                win1.alignChildren = ["fill", "top"];
                win1.spacing = 8;

                var rowWidth = addRow(win1, "Ample: ");
                var inputNum3 = rowWidth.add("edittext", undefined, anchoDeseadoOriginal);
                inputNum3.preferredSize.width = 60;

                var btnGroup = win1.add("group");
                btnGroup.alignment = ["right", "center"];
                btnGroup.add("button", undefined, "D'acord", {name: "ok"});
                btnGroup.add("button", undefined, "Cancel·lar", {name: "cancel"});

                if (win1.show() !== 1) {
                    anchoDeseado = anchoDeseadoOriginal;
                    tempFrame.remove();
                    exit();
                }

                var rawText = inputNum3.text.replace(",", ".");
                anchoDeseado = parseFloat(rawText) || 0;    
            }
            
            imagen.absoluteHorizontalScale = 100;
            imagen.absoluteVerticalScale = 100;

            var boundsOriginales = marco.geometricBounds;
            var anchoOriginal = boundsOriginales[3] - boundsOriginales[1];
            var altoOriginal = boundsOriginales[2] - boundsOriginales[0];
            var proporcion = altoOriginal / anchoOriginal;
            var altoDeseado = anchoDeseado * proporcion;
            
            marco.geometricBounds = [top, left, top + altoDeseado, left + anchoDeseado];
            marco.fit(FitOptions.CONTENT_TO_FRAME);

            var margenPie = 2;   
            var topPie = top + altoDeseado + margenPie;
            var marcoTexto = page.textFrames.add();
            marcoTexto.geometricBounds = [topPie, left, topPie + 10, left + anchoDeseado];

            marcoTexto.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
            marcoTexto.textFramePreferences.verticalJustification = VerticalJustification.TOP_ALIGN;
            marcoTexto.textFramePreferences.insetSpacing = [0, 0, 0, 0];

            if (flag2 && parrafoEncontrado != null) {
                parrafoEncontrado.duplicate(LocationOptions.AT_BEGINNING, marcoTexto.insertionPoints[0]);
                parrafoEncontrado.remove();
            } else {
                marcoTexto.contents = "peu no trobat " + numCrida;
            }

            if (estiloPie.isValid) {
                marcoTexto.paragraphs[0].applyParagraphStyle(estiloPie, false);
            }

            if (marcoTexto.paragraphs.length > 0) {
                var p = marcoTexto.paragraphs[0];
                if (p.characters.length > 0 && p.characters[-1].contents === "\r") {
                    p.characters[-1].remove();
                }
            }
            if (marcoTexto.characters.length > 0 && marcoTexto.characters[-1].contents === "\r") {
                marcoTexto.characters[-1].remove();
            }

            doc.recompose();
            marcoTexto.fit(FitOptions.FRAME_TO_CONTENT);
            var boundsActuales = marcoTexto.geometricBounds; // [top, left, bottom, right]
            var altoCalculado = boundsActuales[2] - boundsActuales[0];
            marcoTexto.geometricBounds = [topPie, left, topPie + altoCalculado, left + anchoDeseado];

            marco.textWrapPreferences.textWrapMode = TextWrapModes.NONE;
            marcoTexto.textWrapPreferences.textWrapMode = TextWrapModes.NONE;

            var grupo = page.groups.add([marco, marcoTexto]);
            grupo.textWrapPreferences.textWrapMode = TextWrapModes.NONE;

            var anchorPage = puntoDeInsercion.parentTextFrames[0].parentPage;
            var anchorX = puntoDeInsercion.horizontalOffset;
            var anchorY = puntoDeInsercion.baseline;
            grupo.move(anchorPage, [anchorX - 35, anchorY - 35]);

            grupo.anchoredObjectSettings.insertAnchoredObject(puntoDeInsercion, AnchorPosition.ANCHORED);

            var settings = grupo.anchoredObjectSettings;
            settings.anchoredPosition = AnchorPosition.ANCHORED;
            settings.anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
            settings.horizontalReferencePoint = AnchoredRelativeTo.ANCHOR_LOCATION;
            settings.verticalReferencePoint = VerticallyRelativeTo.LINE_BASELINE;
            settings.anchorXoffset = -35;
            settings.anchorYoffset = -35;

            doc.recompose();
        }

        tempFrame.remove();
    } else {

        alert("No se han encontrado archivos PNG.");

    }

}
