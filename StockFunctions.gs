function RellenarTabla() {
  // 1. FORZAR ACTUALIZACIÓN DE FÓRMULAS
  SpreadsheetApp.flush(); 

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnaAValidar = 1;       // Columna donde no se puede repetir (1 = Columna A)
  var filaInicioDatos = 20;       // Fila donde empiezan tus datos
  
  // Ahora leemos los valores calculados en tiempo real gracias al flush()
  var producto = hoja.getRange("B5").getValue();
  var stockdisp = hoja.getRange("A11").getValue();
  
  // Validación por si le dan al botón sin haber seleccionado un producto
  if (!producto) {
    SpreadsheetApp.getUi().alert("⚠️ Por favor, selecciona un producto antes de agregarlo.");
    return;
  }

  var existeDuplicado = false;
  var ultimaFilaEfectiva = hoja.getLastRow();
  
  // Definimos exactamente en qué fila vamos a escribir
  var filaDestino = ultimaFilaEfectiva < filaInicioDatos ? filaInicioDatos : ultimaFilaEfectiva + 1;

  // Lógica para no incluir un producto ya existente en la lista
  if (ultimaFilaEfectiva >= filaInicioDatos) {
    var cantidadFilas = ultimaFilaEfectiva - filaInicioDatos + 1;
    var valoresColumna = hoja.getRange(filaInicioDatos, columnaAValidar, cantidadFilas, 1).getValues();
    
    var listaPlana = valoresColumna.map(function(fila) {
      return fila[0].toString().toLowerCase().trim(); 
    });
    
    if (listaPlana.indexOf(producto.toLowerCase().trim()) !== -1) {
      existeDuplicado = true;
    }
  }
  
  // 3. CONDICIONAL: Si ya existe, avisa al usuario; si no, lo inserta
  if (existeDuplicado) {
    SpreadsheetApp.getUi().alert("Error: El valor '" + producto + "' ya existe en la tabla.");
  } else {
    var datos = [producto, "", stockdisp];

    // Si ya es el segundo producto o más (filaDestino > 20), copiamos el formato de la fila anterior
    if (filaDestino > filaInicioDatos){
      var filaOrigen = filaDestino - 1; 
      hoja.insertRowAfter(filaOrigen);
      
      var rangoOrigen = hoja.getRange(filaOrigen, 1, 1, datos.length);
      var rangoDestino = hoja.getRange(filaDestino, 1, 1, datos.length);
      rangoOrigen.copyTo(rangoDestino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    
    // Escribimos los datos en la fila correspondiente
    hoja.getRange(filaDestino, 1, 1, datos.length).setValues([datos]);

    // Opcional: Seleccionar automáticamente la celda de cantidad para agilizar la escritura
    hoja.getRange(filaDestino, 2).activate();
  }
}
//Funcion de borrado de la tabla de productos a vender
function eliminarFilasTablaDejarUna() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 1. CONFIGURACIÓN: Ajusta según tu tabla
  var filaInicioDatos = 20;  // Fila del primer dato (debajo del encabezado)
  var columnaInicio = 1;    // Columna inicial (1 = A)
  var totalColumnas = 4;    // Ancho de la tabla (ej. 5 columnas de la A a la E)
  
  var ultimaFilaHoja = hoja.getLastRow();
  
  // 2. Si hay más de una fila con datos, eliminamos las sobrantes
  if (ultimaFilaHoja > filaInicioDatos)  {
    
    // La eliminación física empieza en la segunda fila de datos
    var filaDondeEmpiezaBorrado = filaInicioDatos + 1;
    
    // Calcula cuántas filas se van a eliminar por completo
    var cantidadFilasAEliminar = ultimaFilaHoja - filaInicioDatos;
    
    // Elimina físicamente las filas de la hoja
    hoja.deleteRows(filaDondeEmpiezaBorrado, cantidadFilasAEliminar);
  }
  // 3. Limpia el contenido de la única fila que quedó para dejarla vacía
  var primeraFilaVacia = hoja.getRange(filaInicioDatos, columnaInicio, 1, totalColumnas);
  primeraFilaVacia.clearContent();
  primeraFilaVacia.removeCheckboxes();
  
  // 4. Deja el cursor seleccionado en esa primera celda
  hoja.getRange(filaInicioDatos, columnaInicio).activate();
}

// 1. Esta función se asigna al botón de la interfaz (se mantiene igual)
function confirmarVenta() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaVenta = libro.getSheetByName("Interfaz Venta"); 
  
  var primeraFilaVenta = 20;
  var ultimaFilaVenta = hojaVenta.getLastRow();
  if (ultimaFilaVenta < 20) {
    SpreadsheetApp.getUi().alert("⚠️ La tabla de ventas está vacía.");
    return;
  }
  
  var rangoVenta = hojaVenta.getRange(primeraFilaVenta, 1, ultimaFilaVenta - primeraFilaVenta + 1, 2); 
  var datosVenta = rangoVenta.getValues();

  // --- VALIDACIÓN DE CAMPOS VACÍOS ---
  var productosVerificados = [];
  for (var i = 0; i < datosVenta.length; i++) {
    var producto = datosVenta[i][0];      
    var cantidadAVender = datosVenta[i][1]; 
    
    if (producto !== "") {
      if (cantidadAVender === "" || isNaN(cantidadAVender) || cantidadAVender <= 0) {
        SpreadsheetApp.getUi().alert(
          "❌ Error: Existen productos que no tienen cantidad a vender, por favor complete estos campos."
        );
        return; 
      }
      productosVerificados.push({producto: producto, cantidad: cantidadAVender});
    }
  }

  if (productosVerificados.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ No hay productos válidos para vender.");
    return;
  }
  
  var htmlTemplate = HtmlService.createTemplateFromFile('ConfirmationWindow.html');
  htmlTemplate.productos = productosVerificados;
  
  var htmlOutput = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(350);
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "📋 Confirmar Resumen de Venta");
}

// 2. NUEVA VERSIÓN BLINDADA: Mantiene los enlaces/URLs de referencia intactos en el Inventario
function ejecutarActualizacionStock() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaVenta = libro.getSheetByName("Interfaz Venta"); 
  var hojaStock = libro.getSheetByName("Inventario"); 
  var hojaHistoricoVentas = libro.getSheetByName("Ventas"); 
  
  var primeraFilaVenta = 20;
  var ultimaFilaVenta = hojaVenta.getLastRow();
  var rangoVenta = hojaVenta.getRange(primeraFilaVenta, 1, ultimaFilaVenta - primeraFilaVenta + 1, 2); 
  var datosVenta = rangoVenta.getValues();
  
  var ultimaFilaStock = hojaStock.getLastRow();
  var rangoStock = hojaStock.getRange(2, 1, ultimaFilaStock - 1, 7);
  
  // === 💡 CAMBIO CLAVE 1: Leemos los datos conservando los enlaces de la columna G ===
  var richTextStock = rangoStock.getRichTextValues();
  var datosStock = rangoStock.getValues(); // Mantenemos este solo para leer los números/textos fácil
  
  var productosNoEncontrados = [];
  var nuevasFilasVentas = [];
  var fechaActual = new Date(); 
  
  for (var i = 0; i < datosVenta.length; i++) {
    var prodVenta = datosVenta[i][0]; 
    var cantVenta = Number(datosVenta[i][1]); 
    
    if (prodVenta === "") continue; 
    var encontrado = false;
    
    for (var j = 0; j < datosStock.length; j++) {
      var prodStock = datosStock[j][2]; // Columna C (Producto)
      
      if (prodVenta === prodStock) {
        var stockActual = Number(datosStock[j][0]); // Columna A (Stock)
        
        // --- VALIDACIÓN: Stock Insuficiente ---
        if (cantVenta > stockActual) {
          SpreadsheetApp.getUi().alert(
            "❌ Error de Stock: La cantidad a vender para '" + prodVenta + 
            "' (" + cantVenta + ") es mayor a la existente en stock (" + stockActual + ")."
          );
          return; 
        }
        
        // Si pasa la validación, procedemos a restar el stock en la matriz normal
        datosStock[j][0] = stockActual - cantVenta;
        
        // Obtener datos adicionales para el registro
        var categoria = datosStock[j][1];      // Columna B (Categoría)
        var precioUnitario = Number(datosStock[j][5]); // Columna F (Precio Unitario)
        
        // Extraemos el objeto link de la celda de referencia para pasarlo al historial con su URL
        var referenciaRich = richTextStock[j][6]; // Columna G (Referencia)
        var linkUrl = referenciaRich.getLinkUrl();
        var textoReferencia = datosStock[j][6];
        
        // Si tiene una URL real adentro, armamos una fórmula hipervínculo para la hoja de Ventas
        var valorReferenciaFinal = linkUrl ? '=HYPERLINK("' + linkUrl + '"; "' + textoReferencia + '")' : textoReferencia;
        
        var total = cantVenta * precioUnitario; 
        
        nuevasFilasVentas.push([
          fechaActual,
          cantVenta,
          categoria,
          prodVenta, 
          precioUnitario,
          valorReferenciaFinal, // Se guarda como enlace funcional también en el histórico de ventas
          total
        ]);
        
        encontrado = true;
        break; 
      }
    }
    
    if (!encontrado) {
      productosNoEncontrados.push(prodVenta);
    }
  }
  
  // === 💡 CAMBIO CLAVE 2: Para no borrar los links del inventario, actualizamos solo el stock (Columna A) ===
  // En lugar de sobreescribir las 7 columnas con texto plano, solo actualizamos los nuevos stocks calculados
  var matrizNuevosStocks = [];
  for (var k = 0; k < datosStock.length; k++) {
    matrizNuevosStocks.push([datosStock[k][0]]);
  }
  hojaStock.getRange(2, 1, datosStock.length, 1).setValues(matrizNuevosStocks);
  
  // Guardamos las nuevas ventas registradas
  if (nuevasFilasVentas.length > 0) {
    var ultimaFilaHistorial = hojaHistoricoVentas.getLastRow();
    hojaHistoricoVentas.getRange(ultimaFilaHistorial + 1, 1, nuevasFilasVentas.length, 7).setValues(nuevasFilasVentas);
  }
  
  if (productosNoEncontrados.length > 0) {
    SpreadsheetApp.getUi().alert("⚠️ Venta procesada, pero los siguientes productos no se encontraron en el Inventario:\n" + productosNoEncontrados.join(", "));
  } else {
    SpreadsheetApp.getUi().alert("✅ ¡Confirmado! El stock ha sido actualizado y la venta ha sido registrada.");
    
    // === 🛠️ LIMPIEZA INTELIGENTE DE LA INTERFAZ ===
    hojaVenta.getRange("B5").clearContent();
    hojaVenta.getRange("A20:C20").clearContent();
    
    var ultimaFilaPostVenta = hojaVenta.getLastRow();
    if (ultimaFilaPostVenta > 20) {
      hojaVenta.deleteRows(21, ultimaFilaPostVenta - 20);
    }
  }
}