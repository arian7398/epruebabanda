
// ID de la hoja de cálculo de Google Sheets
const SPREADSHEET_ID = '1_JdCOrzmf9D-QiMGECX-jzXzR05RpHFbzEIMsvM7hkA';

// Función para mostrar la interfaz web
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Registro de Casos de Extorsión')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Función para guardar datos en el Google Sheet
function guardarDatos(datos) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getActiveSheet();
    
    // Convertir datos complejos a cadenas para almacenamiento
    const agraviadosTexto = datos.agraviados ? 
      datos.agraviados.map(a => a.nombre).join(', ') : '';
    
    const datosInteresAgraviados = datos.datosInteresAgraviados ? 
      datos.datosInteresAgraviados.map(da => 
        `${da.nombre}: Datos Interés - ${da.datosInteres}, Empresa - ${da.tipoEmpresa} ${da.nombreEmpresa}`
      ).join(' | ') : '';
    
    const denunciadosTexto = datos.denunciadosAlias ? 
      datos.denunciadosAlias.map(d => d.nombre || d).join(', ') : '';
    
    const instrumentosTexto = Array.isArray(datos.instrumentosExtorsion) ? 
      datos.instrumentosExtorsion.join(', ') : '';
    
    const numerosTexto = Array.isArray(datos.numerosTelefonicos) ? 
      datos.numerosTelefonicos.join(', ') : '';
    
    const imeiTexto = Array.isArray(datos.imeisTelefonicos) ? 
      datos.imeisTelefonicos.join(', ') : '';
    
    const titularesTexto = Array.isArray(datos.titularesPago) ? 
      datos.titularesPago.join(', ') : '';
    
    sheet.appendRow([
      datos.id,
      new Date(), 
      
      // Información General
      datos.fiscalia,
      datos.fiscalacargo,
      
      // Detalles del Caso
      datos.unidadInteligencia,
      datos.instructorCargo,
      datos.formaInicio,
      datos.carpetaFiscal,
      datos.fechaHecho,
      datos.fechaIngresoFiscal,
      datos.fechadeingreso,
      
      // Información del Agraviado
      datos.tipoAgraviado,
      agraviadosTexto,
      datosInteresAgraviados,
      
      // Información del Denunciado
      datos.delitos,
      datos.lugarHechos,
      denunciadosTexto,
      datos.datosInteresDenunciado,
      datos.nombreBandaCriminal,
      datos.cantidadMiembrosBanda,
      datos.modalidadViolencia,
      datos.modalidadAmenaza,
      datos.atentadosCometidos,
      
      // Instrumentos y Métodos de Extorsión
      instrumentosTexto,
      datos.formaPago,
      numerosTexto,
      imeiTexto,
      datos.cuentaPago,
      titularesTexto,
      
      // Datos de Interés de Pagos
      datos.tipoPago,
      datos.montoSolicitado,
      datos.montoPagado,
      datos.numeroPagos,
      datos.tipoPagoOtros,
      
      // Sumilla y Observaciones
      datos.sumillaHechos,
      datos.observaciones
    ]);
    
    return {
      success: true,
      message: 'Registro guardado exitosamente',
      data: datos
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error al guardar datos: ' + error.toString()
    };
  }
}

// Función para actualizar un registro existente en Google Sheets
function actualizarRegistroEnSheet(datos) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    // Procesar nuevos campos:      
    const formaInicioTexto = datos.formaInicio === 'Otros' ? datos.otrosFormaInicio : datos.formaInicio;

    // Convertir datos complejos a cadenas para almacenamiento
      /*      const agraviadosTexto = datos.agraviados ? 
      datos.agraviados.map(a => a.nombre).join(', ') : ''; */
    // Datos de agraviados con detalles

    const agraviadosTexto = datos.agraviados ? 
      datos.agraviados.map(a => 
        `${a.nombre} (Función: ${a.funcion || 'No especificada'}, Empresa: ${a.tipoEmpresa ? a.tipoEmpresa + ' - ' + a.empresa : 'No especificada'})`
      ).join('\n') : '';
    
    // Bandas criminales
    const bandasCriminalesTexto = Array.isArray(datos.bandasCriminales) ? 
      datos.bandasCriminales.join('\n') : '';

    // Cuentas de pago
    const cuentasPagoTexto = Array.isArray(datos.cuentasPago) ? 
      datos.cuentasPago.join('\n') : '';


    const datosInteresAgraviados = datos.datosInteresAgraviados ? 
      datos.datosInteresAgraviados.map(da => 
        `${da.nombre}: Datos Interés - ${da.datosInteres}, Empresa - ${da.tipoEmpresa} ${da.nombreEmpresa}`
      ).join(' | ') : '';
    
    const denunciadosTexto = datos.denunciadosAlias ? 
      datos.denunciadosAlias.map(d => d.nombre || d).join(', ') : '';
    
    const instrumentosTexto = Array.isArray(datos.instrumentosExtorsion) ? 
      datos.instrumentosExtorsion.join(', ') : '';
    
    const numerosTexto = Array.isArray(datos.numerosTelefonicos) ? 
      datos.numerosTelefonicos.join(', ') : '';
    
    const imeiTexto = Array.isArray(datos.imeisTelefonicos) ? 
      datos.imeisTelefonicos.join(', ') : '';
    
    const titularesTexto = Array.isArray(datos.titularesPago) ? 
      datos.titularesPago.join(', ') : '';
    
    // Buscar la fila que contiene el ID
    let filaEncontrada = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === datos.id) {
        filaEncontrada = i + 1; // +1 porque las filas en Sheets empiezan en 1
        break;
      }
    }

    // Datos de interés del denunciado
    let datosInteresDenunciadoTexto = '';
    if (datos.datosInteresDenunciado) {
      const datosInteres = [];
      if (datos.datosInteresDenunciado.tatuajes) datosInteres.push('Tatuajes');
      if (datos.datosInteresDenunciado.cicatrices) datosInteres.push('Cicatrices');
      if (datos.datosInteresDenunciado.otros) datosInteres.push(datos.datosInteresDenunciado.otrosTexto);
      datosInteresDenunciadoTexto = datosInteres.join('\n');
    }

    
    if (filaEncontrada > 0) {
      // Actualizar los valores en la fila encontrada
      // Las posiciones de columna corresponden a la estructura de guardarDatos
      sheet.getRange(filaEncontrada, 3).setValue(datos.fiscalia);
      sheet.getRange(filaEncontrada, 4).setValue(datos.fiscalacargo);
      sheet.getRange(filaEncontrada, 5).setValue(datos.unidadInteligencia);
      sheet.getRange(filaEncontrada, 6).setValue(datos.instructorCargo);
      sheet.getRange(filaEncontrada, 7).setValue(datos.formaInicio);
      sheet.getRange(filaEncontrada, 8).setValue(datos.carpetaFiscal);
      sheet.getRange(filaEncontrada, 9).setValue(datos.fechaHecho);
      sheet.getRange(filaEncontrada, 10).setValue(datos.fechaIngresoFiscal);
      sheet.getRange(filaEncontrada, 11).setValue(datos.fechadeingreso);
      sheet.getRange(filaEncontrada, 12).setValue(datos.tipoAgraviado);
      sheet.getRange(filaEncontrada, 13).setValue(agraviadosTexto);
      sheet.getRange(filaEncontrada, 14).setValue(datosInteresAgraviados);
      sheet.getRange(filaEncontrada, 15).setValue(datos.delitos);
      sheet.getRange(filaEncontrada, 16).setValue(datos.lugarHechos);
      sheet.getRange(filaEncontrada, 17).setValue(denunciadosTexto);
      sheet.getRange(filaEncontrada, 18).setValue(datos.datosInteresDenunciado);
      sheet.getRange(filaEncontrada, 19).setValue(datos.nombreBandaCriminal);
      sheet.getRange(filaEncontrada, 20).setValue(datos.cantidadMiembrosBanda);
      sheet.getRange(filaEncontrada, 21).setValue(datos.modalidadViolencia);
      sheet.getRange(filaEncontrada, 22).setValue(datos.modalidadAmenaza);
      sheet.getRange(filaEncontrada, 23).setValue(datos.atentadosCometidos);
      sheet.getRange(filaEncontrada, 24).setValue(instrumentosTexto);
      sheet.getRange(filaEncontrada, 25).setValue(datos.formaPago);
      sheet.getRange(filaEncontrada, 26).setValue(numerosTexto);
      sheet.getRange(filaEncontrada, 27).setValue(imeiTexto);
      sheet.getRange(filaEncontrada, 28).setValue(datos.cuentaPago);
      sheet.getRange(filaEncontrada, 29).setValue(titularesTexto);
      sheet.getRange(filaEncontrada, 30).setValue(datos.tipoPago);
      sheet.getRange(filaEncontrada, 31).setValue(datos.montoSolicitado);
      sheet.getRange(filaEncontrada, 32).setValue(datos.montoPagado);
      sheet.getRange(filaEncontrada, 33).setValue(datos.numeroPagos);
      sheet.getRange(filaEncontrada, 34).setValue(datos.tipoPagoOtros);
      sheet.getRange(filaEncontrada, 35).setValue(datos.sumillaHechos);
      sheet.getRange(filaEncontrada, 36).setValue(datos.observaciones);

      return {
        success: true,
        message: 'Registro actualizado exitosamente en Google Sheets',
        data: datos
      };
    } else {
      // No se encontró el registro, guardar como nuevo
      return guardarDatos(datos);
    }
  } catch (error) {
    return {
      success: false,
      message: 'Error al actualizar datos: ' + error.toString()
    };
  }
}



function inicializarSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Hoja de cálculo abierta correctamente: " + ss.getName());
    
    // Buscar específicamente la "Hoja 1" o usar la hoja activa como respaldo
    var sheet = ss.getSheetByName("Hoja 1") || ss.getActiveSheet();
    Logger.log("Usando hoja: " + sheet.getName());
    
    // Verificar el estado actual de la hoja
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    Logger.log("Estado actual: " + numRows + " filas, " + numCols + " columnas");
    
    // Verificar si ya existen encabezados
    var primeraFila = (numRows > 0) ? sheet.getRange(1, 1, 1, 1).getValue() : "";
    var encabezadosExisten = (primeraFila === "ID");

    // Solo crear encabezados si no existen
    if (!encabezadosExisten) {
      Logger.log("No se encontraron encabezados, creándolos ahora");
      
      // Si hay datos pero no son encabezados, insertar una fila al inicio
      if (numRows > 0) {
        Logger.log("Insertando fila al inicio para encabezados");
        sheet.insertRowBefore(1);
      }
    }
    
    // Definir encabezados
    var encabezados = [
      'ID',
      'Fecha de Registro',
      
      // Información General
      'Fiscalía',
      'Fiscal a Cargo',
      
      // Detalles del Caso
      'Unidad de Inteligencia',
      'Instructor a Cargo',
      'Forma de Inicio de Investigación',
      'Carpeta Fiscal',
      'Fecha del Hecho',
      'Fecha Ingreso Carpeta Fiscal',
      'Fecha de Ingreso',
      
      // Información del Agraviado
      'Tipo de Agraviado',
      'Agraviados',
      'Datos de Interés de Agraviados',
      
      // Información del Denunciado
      'Delitos',
      'Lugar de los Hechos',
      'Denunciados',
      'Datos de Interés del Denunciado',
      'Nombre/Apodo Banda Criminal',
      'Cantidad de Miembros de Banda',
      'Modalidad de Violencia',
      'Modalidad de Amenaza',
      'Atentados Cometidos',
      
      // Instrumentos y Métodos de Extorsión
      'Instrumentos de Extorsión',
      'Forma de Pago',
      'Números Telefónicos',
      'IMEI de Teléfonos',
      'Cuenta de Pago',
      'Titulares de Pago',
      
      // Datos de Interés de Pagos
      'Tipo de Pago',
      'Monto Solicitado',
      'Monto Pagado',
      'Número de Pagos',
      'Otros Tipos de Pago',
      
      // Sumilla y Observaciones
      'Sumilla de Hechos',
      'Observaciones'
    ];
    
    // Establecer los valores de encabezado en la primera fila
    sheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    // Dar formato a los encabezados
    sheet.getRange(1, 1, 1, encabezados.length)
        .setBackground("#f3f3f3")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    
    // Congelar la primera fila para que siempre sea visible
    sheet.setFrozenRows(1);
    
    // Ajustar automáticamente el ancho de las columnas
    sheet.autoResizeColumns(1, encabezados.length);
    
    Logger.log("Encabezados creados con éxito");
    
    return {
      success: true,
      message: 'Hoja de cálculo inicializada correctamente'
    };
  } catch (error) {
    Logger.log("ERROR en inicializarSheet: " + error.toString());
    return {
      success: false,
      message: 'Error al inicializar la hoja de cálculo: ' + error.toString()
    };
  }
}

// Función adicional para forzar la creación de encabezados de manera más directa
function crearEncabezadosForzados() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Hoja 1") || ss.getActiveSheet();
    Logger.log("Forzando encabezados en hoja: " + sheet.getName());
    
    // Limpiar la primera fila o insertar una nueva
    if (sheet.getLastRow() > 0) {
      sheet.insertRowBefore(1);
    }
    
    // Lista de encabezados
    var encabezados = [
      'ID',
      'Fecha de Registro',
      'Fiscalía',
      'Fiscal a Cargo',
      'Unidad de Inteligencia',
      'Instructor a Cargo',
      'Forma de Inicio de Investigación',
      'Carpeta Fiscal',
      'Fecha del Hecho',
      'Fecha Ingreso Carpeta Fiscal',
      'Fecha de Ingreso',
      'Tipo de Agraviado',
      'Agraviados',
      'Datos de Interés de Agraviados',
      'Delitos',
      'Lugar de los Hechos',
      'Denunciados',
      'Datos de Interés del Denunciado',
      'Nombre/Apodo Banda Criminal',
      'Cantidad de Miembros de Banda',
      'Modalidad de Violencia',
      'Modalidad de Amenaza',
      'Atentados Cometidos',
      'Instrumentos de Extorsión',
      'Forma de Pago',
      'Números Telefónicos',
      'IMEI de Teléfonos',
      'Cuenta de Pago',
      'Titulares de Pago',
      'Tipo de Pago',
      'Monto Solicitado',
      'Monto Pagado',
      'Número de Pagos',
      'Otros Tipos de Pago',
      'Sumilla de Hechos',
      'Observaciones'
    ];
    
    // Colocar encabezados directamente, uno por uno
    for (var i = 0; i < encabezados.length; i++) {
      sheet.getRange(1, i+1).setValue(encabezados[i]);
    }
    
    // Dar formato
    sheet.getRange(1, 1, 1, encabezados.length)
      .setBackground("#f3f3f3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Congelar la primera fila
    sheet.setFrozenRows(1);
    
    Logger.log("Encabezados forzados creados con éxito");
    return "Encabezados creados con éxito";
  } catch (error) {
    Logger.log("ERROR en crearEncabezadosForzados: " + error.toString());
    return "Error: " + error.toString();
  }
}


function obtenerRegistros() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    Logger.log('Filas encontradas: ' + data.length);

    // Si solo hay encabezados o está vacía, devolver array vacío---------------------------------------
    if (data.length <= 1) {
      return [];
    }
    //---------------------------------------------------------------------

    // Obtener encabezados (primera fila)
    var headers = data[0];
    
    // Convertir datos a objetos
    var registros = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var registro = {
        id: row[0],
        fechaRegistro: row[1],
        
        // Información General
        fiscalia: row[2],
        fiscalacargo: row[3],
        
        // Detalles del Caso
        unidadInteligencia: row[4],
        instructorCargo: row[5],
        formaInicio: row[6],
        carpetaFiscal: row[7],
        fechaHecho: row[8],
        fechaIngresoFiscal: row[9],
        fechadeingreso: row[10],
        
        // Información del Agraviado
        tipoAgraviado: row[11],

        agraviados: (row[12] && typeof row[12] === 'string') ? 
          row[12].split(', ').map(function(nombre) {
          return {id: Utilities.getUuid(), nombre: nombre.trim()};
        }) : [],
        
        // Información del Denunciado
        delitos: row[14],
        lugarHechos: row[15],

        // Para denunciados:
        denunciadosAlias: (row[16] && typeof row[16] === 'string') ? 
          row[16].split(', ').map(function(nombre) {
          return {id: Utilities.getUuid(), nombre: nombre.trim()};
        }) : [],

        datosInteresDenunciado: row[17],
        nombreBandaCriminal: row[18],
        cantidadMiembrosBanda: row[19],
        modalidadViolencia: row[20],
        modalidadAmenaza: row[21],
        atentadosCometidos: row[22],

        // Instrumentos y Métodos de Extorsión
        instrumentosExtorsion: (row[23] && typeof row[23] === 'string') ? 
          row[23].split(', ').map(function(s) { return s.trim(); }) : [],
        formaPago: row[24],
        numerosTelefonicos: (row[25] && typeof row[25] === 'string') ? 
          row[25].split(', ').map(function(s) { return s.trim(); }) : [],
        imeisTelefonicos: (row[26] && typeof row[26] === 'string') ? 
          row[26].split(', ').map(function(s) { return s.trim(); }) : [],
        cuentaPago: row[27],
        titularesPago: (row[28] && typeof row[28] === 'string') ? 
         row[28].split(', ').map(function(s) { return s.trim(); }) : [],
      
        /// Datos de Interés de Pagos
        tipoPago: row[29],
        montoSolicitado: row[30],
        montoPagado: row[31],
        numeroPagos: row[32],
        tipoPagoOtros: row[33], 
        
        // Sumilla y Observaciones
        sumillaHechos: row[34],
        observaciones: row[35]
      };

      Logger.log('Registros procesados: ' + registros.length);
      registros.push(registro);
    }
    
    return registros;
  } catch (error) {
    Logger.log('Error en obtenerRegistros: ' + error);
    throw error;
  }
}


// Función para eliminar un registro
function eliminarRegistro(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    // Buscar la fila que contiene el ID
    let filaEncontrada = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        filaEncontrada = i + 1; // +1 porque las filas en Sheets empiezan en 1
        break;
      }
    }
    
    if (filaEncontrada > 0) {
      // Eliminar la fila
      sheet.deleteRow(filaEncontrada);
      
      return {
        success: true,
        message: 'Registro eliminado correctamente'
      };
    } else {
      return {
        success: false,
        message: 'No se encontró el registro para eliminar'
      };
    }
  } catch (error) {
    return {
      success: false,
      message: 'Error al eliminar el registro: ' + error.toString()
    };
  }
}
