var doc = app.activeDocument;

//var carpeta = Folder.selectDialog(
//    "Selecciona la carpeta on hi ha les fotos"
//);

var carpeta = Folder("C:/Users/jordi/Desktop/provaScript/figures2")
var archivoTxt = File("C:/Users/jordi/Desktop/provaScript/peus_de_fotos.txt");


var primeraCrida = 0; // i-1
var ultimaCrida = 4 ; // i

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

if (carpeta != null) {
    var fotos = carpeta.getFiles("*.png"); // fotos es un array

    archivoTxt.open("r");
    var contenido = archivoTxt.read();
    archivoTxt.close();
    var listaPies = contenido.split(/\r?\n/);
    

    if (fotos.length > 0) {

        // i index crida, j index foto
        for (var i = ultimaCrida - 1; i >= primeraCrida; i--) {
            //var page = doc.pages[0];
            var nombreEstilo = "Peu";
            var estiloPie = doc.paragraphStyles.itemByName(nombreEstilo);
            var textBuscar = "(fig. " + (i + 1) + ")";
            var top = 40;
            var left = 40;

            var numCrida = i+1;


            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.findWhat = textBuscar;
            var resultados = doc.findText();
            if (resultados.length === 0) {
                alert("no trobat el text")
                continue;
            }

            var resultadoEncontrado = resultados[0];
            var puntoDeInsercion = resultadoEncontrado.insertionPoints[-1];
            var page = puntoDeInsercion.parentTextFrames[0].parentPage;

            var flag1 = false;
            var j = fotos.length -1;
            while (j >= 0 && !flag1){

                var foto = fotos[j];
                var nomFoto = String(decodeURIComponent(foto.name));
                var numFoto = extraerNumero(nomFoto)

                if (numFoto == numCrida){
                    flag1 = true;
                } else {
                    j--;
                }
            }
            if (!flag1) {
                alert("foto no trobada")
                continue;
            }

            var flag2 = false;
            var j = listaPies.length -1;
            while (j >= 0 && !flag2){

                var varPeuFoto = listaPies[j];
                var numPeu = extraerNumero(varPeuFoto)

                if (numPeu == numCrida){
                    flag2 = true;
                } else {
                    j--;
                }
            }
            if (!flag2) {
                alert("peu no trobat")
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
            

            marcoTexto.contents = varPeuFoto + " ";
            if (estiloPie.isValid) {
                marcoTexto.paragraphs[0].applyParagraphStyle(estiloPie, true);
            }

            // Opcional: autoajustar el alto de la caja de texto al contenido escrito
            marcoTexto.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
            marcoTexto.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;

            var grupo = page.groups.add([marco, marcoTexto]);
            //grupo.textWrapPreferences.textWrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
            
            var anchorPage = puntoDeInsercion.parentTextFrames[0].parentPage;
            var anchorX = puntoDeInsercion.horizontalOffset;
            var anchorY = puntoDeInsercion.baseline;
            grupo.move(anchorPage, [anchorX -35, anchorY - 35]);
            grupo.anchoredObjectSettings.insertAnchoredObject(puntoDeInsercion, AnchorPosition.ANCHORED);

        }

    } else {

        alert("No se han encontrado archivos PNG.");

    }

}
