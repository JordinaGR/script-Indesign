var doc = app.activeDocument;

var carpeta = Folder.selectDialog("Selecciona la carpeta on hi ha les fotos");
var archivoWord = File.openDialog("Selecciona el document de Word amb els peus de foto", "Archivos de Word:*.docx;*.doc;");

//var carpeta = Folder("C:/Users/jordi/Desktop/provaScript/figures2")
//var archivoTxt = File("C:/Users/jordi/Desktop/provaScript/peus_de_fotos.txt");
//var archivoWord = File("C:/Users/jordi/Desktop/provaScript/peus.docx");

if (!carpeta || !archivoWord) {
    alert("Operació cancel·lada: Cal seleccionar una carpeta i un fitxer.");
    exit();
}

// 2. Finestra simplificada (ScriptUI) per als 2 strings i 2 enters
var win = new Window("dialog", "Paràmetres d'entrada");
win.alignChildren = "left";
win.spacing = 8;

// Camps de text
win.add("statictext", undefined, "format");
var inputStr1 = win.add("edittext", undefined, "Fig ");
inputStr1.preferredSize.width = 120;

var inputStr2 = win.add("edittext", undefined, "");
inputStr2.preferredSize.width = 120;

// Camps numèrics
win.add("statictext", undefined, "num. primera figura");
var inputNum1 = win.add("edittext", undefined, "1");
inputNum1.preferredSize.width = 120;

win.add("statictext", undefined, "num. última figura");
var inputNum2 = win.add("edittext", undefined, "10");
inputNum2.preferredSize.width = 120;

win.add("statictext", undefined, "nom de l'estil de paràgraf");
var inputStr3 = win.add("edittext", undefined, "Peu");
inputStr3.preferredSize.width = 120;

// Botons
var btnGroup = win.add("group");
btnGroup.add("button", undefined, "D'acord", {name: "ok"});
btnGroup.add("button", undefined, "Cancel·lar", {name: "cancel"});

// Mostrar la finestra
if (win.show() !== 1) {
    exit(); // L'usuari ha tancat o premut Cancel·lar
}

// Guardem els valors introduïts
var pre = inputStr1.text;
var post = inputStr2.text;
var nombreEstilo = inputStr3.text;

// Convertim el text dels enters a números (parseInt)
var primeraCrida = parseInt(inputNum1.text, 10) || 0;
var ultimaCrida = parseInt(inputNum2.text, 10) || 0;

// primeraCrida = 0; // i-1
//var ultimaCrida = 4 ; // i
//var pre = "(fig. ";
//var post = ")";

var myExtensions = [".png", ".jpg", ".jpeg", ".tif", ".tiff", ".gif", ".pdf", ".psd", ".ai", ".heic"];


function extraerNumero(texto) {
    if (!texto) return null;
    
    // Forzamos la conversión a cadena primitiva de JavaScript
    var str = String(texto) + "";
    var digitosEncontrados = "";
    var leyendoNumero = false;

    for (var i = 0; i < str.length; i++) {
        // En ExtendScript es más seguro acceder directamente con corchetes [i]
        var cchar = str[i];
        
        // Comprobar si el carácter actual es un dígito entre '0' y '9'
        if (cchar >= '0' && cchar <= '9') {
            digitosEncontrados += cchar;
            leyendoNumero = true;
        } else if (leyendoNumero) {
            // Si ya estábamos leyendo un número y encontramos un no-dígito, terminamos
            break;
        }
    }

    // Si encontramos dígitos, los convertimos a un número entero
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
    // Comprovar per tipus de fitxer intern de Mac
    switch (myFileType) {
        case "JPEG":
        case "PNGf":
        case "TIFF":
            return true;
        default:
            // Comprovar per extensió en el nom de fitxer
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
    //var listaPies = obtenerPiesDesdeWord(archivoWord, doc);

    if (fotos.length > 0) {
        
        var tempFrame = doc.pages[0].textFrames.add();
        tempFrame.geometricBounds = [0, 0, 200, 200];
        tempFrame.place(archivoWord);
        
        // i index crida, j index foto
        for (var numCrida = primeraCrida; numCrida <= ultimaCrida; numCrida++) {
            //var page = doc.pages[0];
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
                alert("no s'ha trobat el text de la crida " + textBuscar)
                continue;
            }

            var resultadoEncontrado = resultados[0];
            var puntoDeInsercion = resultadoEncontrado.insertionPoints[-1];
            //var page = puntoDeInsercion.parentTextFrames[0].parentPage;

            var parentFrames = puntoDeInsercion.parentTextFrames;
            
            // Check if parentTextFrames exists and has at least one frame
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
                    parrafoEncontrado = parrafosTemp[k]; // Reference to the actual formatted InDesign paragraph
                } else {
                    k++;
                }
            }
            if (!flag2) {
                alert("peu no trobat " + numCrida)
                continue;
            }

            var colocados = page.place(foto, [left, top]);
            var imagen = colocados[0];        // El objeto de la imagen importada
            var marco = imagen.parent;        // El marco contenedor de la imagen

            var anchoDeseado = 80; // en pt ample de les fotos
            
            imagen.absoluteHorizontalScale = 100;
            imagen.absoluteVerticalScale = 100;


            // 3. Calculamos la proporción original de la imagen
            var boundsOriginales = marco.geometricBounds;
            var anchoOriginal = boundsOriginales[3] - boundsOriginales[1];
            var altoOriginal = boundsOriginales[2] - boundsOriginales[0];
            var proporcion = altoOriginal / anchoOriginal;

            // 4. Calculamos el alto proporcional exacto según el ancho deseado
            var altoDeseado = anchoDeseado * proporcion;

            // 5. Aplicamos las dimensiones EXACTAS al marco contenedor
            marco.geometricBounds = [top, left, top + altoDeseado, left + anchoDeseado];
            // 6. Forzamos a la imagen a rellenar el marco proporcionalmente (esto SÍ la amplía)
            marco.fit(FitOptions.CONTENT_TO_FRAME);


            // 3. Crear el marco de texto justo debajo de la imagen
            var margenPie = 2;   
            var altoPie = 5;
            
            var topPie = top + altoDeseado + margenPie;
            var bottomPie = topPie + altoPie;
            var rightPie = left + anchoDeseado;

            var marcoTexto = page.textFrames.add();
            // Posicionamos el marco de texto [top, left, bottom, right]
            marcoTexto.geometricBounds = [topPie, left, bottomPie, rightPie];
            
            parrafoEncontrado.duplicate(LocationOptions.AT_BEGINNING, marcoTexto.insertionPoints[0]);
            parrafoEncontrado.remove();

            if (estiloPie.isValid) {
                marcoTexto.paragraphs[0].applyParagraphStyle(estiloPie, false);
            }

            // 1. Disable text wrapping on frames before grouping
            marco.textWrapPreferences.textWrapMode = TextWrapModes.NONE;
            marcoTexto.textWrapPreferences.textWrapMode = TextWrapModes.NONE;

            // 2. Create the group
            var grupo = page.groups.add([marco, marcoTexto]);
            grupo.textWrapPreferences.textWrapMode = TextWrapModes.NONE;

            // 3. Move group near target location
            var anchorPage = puntoDeInsercion.parentTextFrames[0].parentPage;
            var anchorX = puntoDeInsercion.horizontalOffset;
            var anchorY = puntoDeInsercion.baseline;
            grupo.move(anchorPage, [anchorX - 35, anchorY - 35]);

            // 4. INSERT FIRST (This enables anchored settings mode)
            grupo.anchoredObjectSettings.insertAnchoredObject(puntoDeInsercion, AnchorPosition.ANCHORED);

            // 5. CONFIGURE / ADJUST AFTER (Maintains standard manual release capabilities)
            var settings = grupo.anchoredObjectSettings;
            settings.anchoredPosition = AnchorPosition.ANCHORED;
            settings.anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
            settings.horizontalReferencePoint = AnchoredRelativeTo.ANCHOR_LOCATION;
            settings.verticalReferencePoint = VerticallyRelativeTo.LINE_BASELINE;
            settings.anchorXoffset = -35;
            settings.anchorYoffset = -35;

            // 6. Force recomposition for next iteration
            doc.recompose();
        }

        //tempFrame.remove();
    } else {

        alert("No se han encontrado archivos PNG.");

    }

}
