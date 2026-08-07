var doc = app.activeDocument;

//var carpeta = Folder.selectDialog(
//    "Selecciona la carpeta on hi ha les fotos"
//);

var carpeta = Folder("C:/Users/jordi/Desktop/provaScript/figures2")
var archivoTxt = File("C:/Users/jordi/Desktop/provaScript/peus_de_fotos.txt");

if (carpeta != null) {
    var fotos = carpeta.getFiles("*.png"); // fotos es un array

    archivoTxt.open("r");
    var contenido = archivoTxt.read();
    archivoTxt.close();
    var listaPies = contenido.split(/\r?\n/);
    
    
    if (fotos.length > 0) {

        for (var i = 0; i < fotos.length; i++){
            var page = doc.pages[0];
            var foto = fotos[i];
            var nombreEstilo = "Peu";
            var estiloPie = doc.paragraphStyles.itemByName(nombreEstilo);
            var textBuscar = "(fig. " + (i + 1) + ")";
            var top = 40;
            var left = 40;



            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.findWhat = textBuscar;
            var resultados = doc.findText();

            if (resultados.length > 0) {
                var resultadoEncontrado = resultados[0];
                var puntoDeInsercion = resultadoEncontrado.insertionPoints[-1];

                var colocados = page.place(foto, [left, top]);
                var imagen = colocados[0];        // El objeto de la imagen importada
                var marco = imagen.parent;        // El marco contenedor de la imagen

                var anchoDeseado = 50; // en pt
                
                var bounds = marco.geometricBounds;
                var anchoActual = bounds[3] - bounds[1];
                var altoActual = bounds[2] - bounds[0];
                var proporcion = altoActual / anchoActual;
                var altoDeseado = anchoDeseado * proporcion;

                marco.geometricBounds = [top, left, top + altoDeseado, left + anchoDeseado];
                
                marco.fit(FitOptions.CONTENT_TO_FRAME);
                marco.textWrapPreferences.textWrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
                marco.textWrapPreferences.textWrapOffset = ["0pt", "0pt", "0pt", "0pt"];


                // 3. Crear el marco de texto justo debajo de la imagen
                var margenPie = 2;   
                var altoPie = 5;
                
                var topPie = top + altoDeseado + margenPie;
                var bottomPie = topPie + altoPie;
                var rightPie = left + anchoDeseado;

                var marcoTexto = page.textFrames.add();
                // Posicionamos el marco de texto [top, left, bottom, right]
                marcoTexto.geometricBounds = [topPie, left, bottomPie, rightPie];
                
                var varPeuFoto = listaPies[i];
                marcoTexto.contents = varPeuFoto + " ";  // canviar els peus de foto
                if (estiloPie.isValid) {
                    marcoTexto.paragraphs[0].applyParagraphStyle(estiloPie, true);
                }

                // Opcional: autoajustar el alto de la caja de texto al contenido escrito
                marcoTexto.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
                marcoTexto.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;

                var grupo = page.groups.add([marco, marcoTexto]);
                grupo.textWrapPreferences.textWrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
                grupo.anchoredObjectSettings.insertAnchoredObject(puntoDeInsercion, AnchorPosition.ANCHORED);

            }
        }

    } else {

        alert("No se han encontrado archivos PNG.");

    }

}