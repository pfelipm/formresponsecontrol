// Ámbito de autorización
/**
 * @OnlyCurrentDoc
 */


function onInstall(e) {
  
  // Otras cosas que se deben hacer siempre
  onOpen(e);
}

function onOpen() {

  // Crear menú de la aplicación
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('✅ Configurar', 'configurar')
    .addItem('❔ Comprobar estado', 'comprobarEstado')
    .addSeparator()
    .addItem('⬇️ Forzar copia de formato', 'formatoForzado')
    .addItem('⬇️ Forzar copia de fórmulas', 'formulasForzado')
    .addItem('⬇️ Forzar copia de validación', 'validacionForzado')
    .addSeparator()
    .addItem('⚠️ Restaurar FRC', 'restaurar')
    .addSeparator()
    .addItem('💡 Sobre FRC', 'acercaDe') 
    .addToUi();
}

function acercaDe() {

  // Presentación del complemento
  var panel = HtmlService.createHtmlOutputFromFile('acercaDe')
    .setWidth(420)
    .setHeight(220)
  SpreadsheetApp.getUi().showModalDialog(panel, '💡 ¿Qué es FRC?');
}

function extenderFormato(filaModelo, filaRespuesta, reprocesar) {

  // Aplica el formato (+ altura + formato condicional) de la fila que se pasa como parámetro
  // a todas por debajo de ella (reprocesar = true) o solo a la última (reprocesar = false);
  // filaRespuesta contiene la correspondiente a la respuesta de formulario que se debe
  // procesar o 0 si se trata de una aplicación manual.
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
    
  // ¿Hay respuestas?
  if (lastRow > filaModelo) {    
  
    // ¿En todas las filas o solo la última?
    if (reprocesar == true) {
      
      // Aplicamos sobre toda la hoja por debajo de "filaModelo"
      // Si se trata de una respuesta previa modificada, mismo tratamiento.
      
      // Formato
      sheet.getRange(filaModelo, 1, 1, lastColumn).copyFormatToRange(sheet, 1, lastColumn, filaModelo + 1, lastRow);
      
      // Altura de fila
      var alturaFila = sheet.getRowHeight(filaModelo);
      sheet.setRowHeights(filaModelo + 1, lastRow - filaModelo, alturaFila);   
    }
    else {
     
     // Aplicamos solo sobre la fila de la respuesta recibida
      
      // Fomato
      sheet.getRange(filaModelo, 1, 1, lastColumn).copyFormatToRange(sheet, 1, lastColumn, filaRespuesta, filaRespuesta);
      
      // Altura de fila
      var alturaFila = sheet.getRowHeight(filaModelo);
      sheet.setRowHeight(filaRespuesta, alturaFila);   
    }
  }   
}

function extenderFormulas(filaModelo, filaRespuesta, reprocesar) {

  // Copia las fórmulas presentes en la fila que se pasa como parámetro
  // a todas por debajo de ella (reprocesar = true) o solo a la última (reprocesar = false)
  // filaRespuesta contiene la correspondiente a la respuesta de formulario que se debe
  // procesar o 0 si se trata de una aplicación manual.
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
    
  // ¿Hay respuestas?
  if (lastRow > filaModelo) {
    
    // Recorremos fila modelo buscando fórmulas
    for (var col = 1; col <= lastColumn; col++) {
      celdaFormula = sheet.getRange(filaModelo,col,1,1);
      
      // Si en alguna celda de la fila 2 hay una fórmula la copiamos donde corresponda
      if (celdaFormula.getFormula() != '') {
        
        // ¿En todas las filas o solo la última?
        if (reprocesar == true) {
          
          // Copiar en todas las filas por debajo
          // Si se trata de una respuesta previa modificada, mismo tratamiento.
          celdaFormula.copyTo(sheet.getRange(filaModelo + 1, col, lastRow - filaModelo, 1));
        }
        else {
          
          // Copiar en la fila de la respuesta recibida
          celdaFormula.copyTo(sheet.getRange(filaRespuesta,col));
        }
      }
    }
  }
}

function extenderValidacion(filaModelo, filaRespuesta, reprocesar) {

  // Copia los ajustes de validación en las celdas de la fila que se pasa como parámetro
  // a todas por debajo de ella (reprocesar = true) o solo a la última (reprocesar = false)
  // filaRespuesta contiene la correspondiente a la respuesta de formulario que se debe
  // procesar o 0 si se trata de una aplicación manual.
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
    
  // ¿Hay respuestas?
  if (lastRow > filaModelo) {
    
    // ¿En todas las filas o solo la última?
    if (reprocesar == true) {
        
      // Aplicar en todas las filas por debajo
      // Si se trata de una respuesta previa modificada, mismo tratamiento.
      sheet.getRange(filaModelo, 1, 1, lastColumn).copyTo(sheet.getRange(filaModelo + 1, 1, lastRow - filaModelo, lastColumn),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    else {
          
      // Copiar en la fila de la respuesta recibida
      sheet.getRange(filaModelo, 1, 1, lastColumn).copyTo(sheet.getRange(filaRespuesta, 1),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
     }
   }
}  

function formatoForzado() {
  
  // El operador "+" convierte la cadena a número
  var fila = +PropertiesService.getDocumentProperties().getProperty('fila');
  
  // ¿Seguimos?
  if (SpreadsheetApp.getUi().alert('¿Deseas continuar?', 'Esta función aplicará el formato de la fila\n\n'
    + '# ' + fila + ' #\n\n a todas las que quedan por debajo.'
    , SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK) {
       
    // Mensaje de inicio de proceso.
    SpreadsheetApp.getActiveSpreadsheet().toast('Aplicando formato...');
    
    extenderFormato(fila, 0, true);
    
    // Mensaje de fin de proceso
    SpreadsheetApp.getActiveSpreadsheet().toast('Formato aplicado.');
  }
}

function formulasForzado() {
  
  var fila = +PropertiesService.getDocumentProperties().getProperty('fila');
  
  // ¿Seguimos?
  if (SpreadsheetApp.getUi().alert('¿Deseas continuar?', 'Esta función copiará las fórmulas de la fila\n\n'
    + '# ' + fila + ' #\n\n a todas las que quedan por debajo.'
    , SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK) {
      
    // Mensaje de inicio de proceso.
    SpreadsheetApp.getActiveSpreadsheet().toast('Copiando fórmulas...');
   
    extenderFormulas(fila, 0, true);
    
    // Mensaje de fin de proceso
    SpreadsheetApp.getActiveSpreadsheet().toast('Fórmulas copiadas.');
  }  
}

function validacionForzado() {
  
  var fila = +PropertiesService.getDocumentProperties().getProperty('fila');
  
  // ¿Seguimos?
  if (SpreadsheetApp.getUi().alert('¿Deseas continuar?', 'Esta función copia los ajustes de validación de datos de la fila\n\n'
    + '# ' + fila + ' #\n\n a todas las que quedan por debajo.'
    , SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK) {
    
     // Mensaje de inicio de proceso.
    SpreadsheetApp.getActiveSpreadsheet().toast('Aplicando validación de datos...');
   
    extenderValidacion(fila, 0, true);}
    
    // Mensaje de fin de proceso
    SpreadsheetApp.getActiveSpreadsheet().toast('Validación aplicada.');
}

function restaurar() {

  // ¿Seguimos?
  if (SpreadsheetApp.getUi().alert('¿Deseas restaurar FRC?',
    '¡PRECAUCIÓN!\n\n' +
    'Esta función *solo* debe utilizarse si el complemento se comporta de modo\n' +
    'errático al procesar en segundo plano las respuestas del formulario y/o \n' +
    'el interruptor del panel de configuración no muestra correctamente su estado \n' +
    'de activación.\n\n' +
    '¡Se restaurarán todos los ajustes por defecto y se desactivará FRC ❌!\n\n' +
    'El procedimiento es más efectivo si TODOS los usuarios con acceso en edición\n' +
    'al documento utilizan esta función en caso de problemas.'
    ,SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK) {
    
    // Localizar y eliminar trigger de tipo respuestas de formulario
    var triggers = ScriptApp.getProjectTriggers();  
    try {
      var ntriggers = 0;
      var triggers = ScriptApp.getProjectTriggers();
      for (i in triggers) {
        if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
          ntriggers++;
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
    }    
    catch (e) {SpreadsheetApp.getUi().alert('¡Error!','Se han producido errores al tratar de desactivar FRC.\n\n'+e,SpreadsheetApp.getUi().ButtonSet.OK);}
    
    // Restaura valores por defecto
    PropertiesService.getDocumentProperties().setProperties({
      'fila' : '2',
      'autoFormato' : 'true',
      'autoFormula' : 'true',
      'autoValidacion' : 'true',
      'autoInversion' : 'false',
      'reprocesar' : 'false',
      'triggerDe' : '',
    }, true);
    
    // Este modo de inicialización parece dar más problemas de sincronización
    /*PropertiesService.getDocumentProperties().deleteAllProperties();
    configurar();  */
    
    SpreadsheetApp.getUi().alert('FRC restaurado', 'Activadores encontrados: ' + '# ' + ntriggers + ' #', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function comprobarEstado() {

  var triggerDe = PropertiesService.getDocumentProperties().getProperty('triggerDe');
  if (triggerDe == '' || triggerDe == null) {
      SpreadsheetApp.getUi().alert('💡 Form Response Control no está activado.');
  }
  else {
    SpreadsheetApp.getUi().alert('💡 Form Response Control ha sido activado por:\n\n' + triggerDe); 
  }
}

function acercaDe() {

  // Presentación del complemento
  var panel = HtmlService.createHtmlOutputFromFile('acercaDe')
    .setWidth(420)
    .setHeight(220)
  SpreadsheetApp.getUi().showModalDialog(panel, '💡 ¿Qué es FRC?');
}

function modificarEstadoFrc(comando) {
  
  // parámetro comando = 'activar' >> activar (instalar trigger) | 'desactivar' >> desactivar (desinstalar trigger)
  // return estado = 'desactivado' | 'activado' | 'invariante'
  
  var estado = '';
  
  // Objeto semáforo
  // El bloqueo es único para cualquier objeto Lock en cualquier
  // parte del código !!
  var semaforo = LockService.getDocumentLock();
  
  try { 
  
    // Queremos fallar inmediatamente
    semaforo.waitLock(1);
    
    // ¿Activar o desactivar FRC?
    
    switch (comando) {
      
      case 'activar': // *** Proceso de activación de FRC ***
    
        // Comprobar si un editor de la hdc ya ha instalado el trigger
        // La identificación del propietario del trigger debe realizarse dentro del bloque de código protegido
        var triggerDe = PropertiesService.getDocumentProperties().getProperty('triggerDe');
        if (triggerDe != '' && triggerDe != null) {
        
          // >>>> Fin de sección de código en exclusión mutua (activado OK) <<<<          
          semaforo.releaseLock(); 
          SpreadsheetApp.getUi().alert('💡 Form Response Control ya ha sido activado por:\n\n' + triggerDe + '\n\n¡Nada que hacer!');
          estado = 'activado';
        }
        else {
          
          // Vamos con ello
          // Interceptar evento de recepción de formulario
          var triggers = ScriptApp.getProjectTriggers();  
          try {
            
            // Instalamos el manejador de onFormSubmit()
            ScriptApp.newTrigger('nuevaRespuestaForm')
            .forSpreadsheet(SpreadsheetApp.getActive())
            .onFormSubmit()
            .create();
            
            // Anotamos qué usuario ha instalado el trigger para tratar de evitar duplicidades
            // ¡No es posible controlar la presencia de triggers instalados manualmente por otros usuarios!
            
            // >>>> Fin de sección de código en exclusión mutua (activado OK) <<<<
            semaforo.releaseLock();      
            
            PropertiesService.getDocumentProperties().setProperty('triggerDe',Session.getEffectiveUser());  
            SpreadsheetApp.getUi().alert('🚀 Form Response Control está activado.');
            estado = 'activado';
            // >>>> Fin de sección de código en exclusión mutua (activado OK) <<<<
            // Se libera Lock tras Alert() para minimizar problemas de actualización de propiedad 'triggerDe' ¿necesario?
            //semaforo.releaseLock();      
          
          } catch (e) { // captura excepción de activación
          
            // >>>> Fin de sección de código en exclusión mutua (error general) <<<<
            semaforo.releaseLock();
            SpreadsheetApp.getUi().alert('¡Error!','Se han producido errores activando el complemento, es posible que no funcione correctamente.\n\n'+e,SpreadsheetApp.getUi().ButtonSet.OK);
            estado = 'desactivado';
          }
        }
        break;
      
      case 'desactivar': // *** Proceso de desactivación de FRC ***

        // Comprobar si el trigger no está instalado
        // La identificación del propietario del trigger debe realizarse dentro del bloque de código protegido  
        var triggerDe = PropertiesService.getDocumentProperties().getProperty('triggerDe');
        if (triggerDe == '' || triggerDe == null) {
        
          // >>>> Fin de sección de código en exclusión mutua (no activado) <<<<
          semaforo.releaseLock(); 
          SpreadsheetApp.getUi().alert('💡 Form Response Control no está activado.\n\n¡Nada que hacer!');
          estado = 'desactivado';
        }
        else {
          
          // Si lo ha instalado el usuario actual, localizar
          if (triggerDe == Session.getEffectiveUser()) {
            var triggerActual = null;
            var triggers = ScriptApp.getProjectTriggers();
            for (var i = 0; i < triggers.length; i++) {
              if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
                triggerActual = triggers[i];
              break;
              }
            }      
            try {
                
              // Eliminar trigger
              ScriptApp.deleteTrigger(triggerActual);
             
              // Actualizar registro de propietario del trigger
              // Se asume que la lógica del código impide activaciones múltiples
              // Por parte de diferentes usuarios
              PropertiesService.getDocumentProperties().setProperty('triggerDe','');
             
              // >>>> Fin de sección de código en exclusión mutua, camino B (desactivado OK) <<<<
              semaforo.releaseLock(); 
    
              SpreadsheetApp.getUi().alert('🛑 Form Response Control está desactivado.');
              estado = 'desactivado';
            }
            catch (e) { // captura excepción de desactivación
              // >>>> Fin de sección de código en exclusión mutua (error general) <<<<
              semaforo.releaseLock(); 
              SpreadsheetApp.getUi().alert('¡Error!','Se han producido errores desactivando el complemento, es posible que no funcione correctamente.\n\n'+e,SpreadsheetApp.getUi().ButtonSet.OK);}     
          }
          else {
            // Solo queda un caso, otro usuario ha activado el trigger, no podemos hacer nada
            // >>>> Fin de sección de código en exclusión mutua <<<<
            semaforo.releaseLock(); 
            SpreadsheetApp.getUi().alert('💡 Form Response Control ha sido activado por:\n\n' + triggerDe + '\n\n¡Pídele que lo desactive!');
            estado = 'activado';
          }
        }
        break;
    }
  }
  catch (e) { // captura excepción de acceso al semáforo
    
    // >>>> Fin de sección de código en exclusión mutua (bloqueado por semáforo) <<<<
    semaforo.releaseLock();
    SpreadsheetApp.getUi().alert('¡Error!','Otro usuario ya está intentado activar o desactivar FRC.\n\n' +
                                 '¡Verifica el estado de activación tras cerrar esta alerta!',SpreadsheetApp.getUi().ButtonSet.OK);
    estado = 'invariante';
  }
 
  return estado;
}

function configurar() {

  // Si es la 1ª vez que se ejecuta, inicializar ajustes (propiedades del documento)
  // Se activan determinadas opciones por defecto (fila 2, extender formato, fórmulas y validación
  if (PropertiesService.getDocumentProperties().getProperty('triggerDe') == null) {
    PropertiesService.getDocumentProperties().setProperties({
      'fila' : '2',
      'autoFormato' : 'true',
      'autoFormula' : 'true',
      'autoValidacion' : 'true',
      'autoInversion' : 'false',
      'reprocesar' : 'false',
      'triggerDe' : '',
    }, true);
  }

  // Script ya configurado, abrimos el panel de configuración
  var panel = HtmlService.createHtmlOutputFromFile('panelAjustes')
    .setHeight(450)
    .setWidth(320);
   // .setTitle('✅ Configuración FRC');
  //ui.showSidebar(panel);  
  SpreadsheetApp.getUi().showModalDialog(panel,'✅ Configuración FRC');  
}

function obtenerPreferencias(){

  // Obtener preferencias guardadas y pasárselas a la interfaz
  return PropertiesService.getDocumentProperties().getProperties();
}

function actualizarPreferencias(preferencias) {

  var propiedadesDocumento = PropertiesService.getDocumentProperties();
  
  // Almacenar ajustes en propiedades del documento para que sean persistentes
  for (var ajuste in preferencias) {
    propiedadesDocumento.setProperty(preferencias[ajuste].clave, preferencias[ajuste].valor.toString());
  }
}

function nuevaRespuestaForm(e) {

  // Aquí está la fiesta...
  
  // Primero comprobemos si disponemos de los permisos necesarios 
  // Tomado de aquí https://developers.google.com/gsuite/add-ons/concepts/triggers#authorizing_installable_triggers
  var addonTitle = 'Form Response Control';
  var props = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger requires authorization that has not
  // been granted yet; if so, warn the user via email. This check is required
  // when using triggers with add-ons to maintain functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to re-authorize; the normal trigger action is not
    // conducted, since it requires authorization first. Send at most one
    // "Authorization Required" email per day to avoid spamming users.
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('emailReAutorizacion');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        var message = html.evaluate();
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
            'Autorización necesaria',
            message.getContent(), {
                name: addonTitle,
                htmlBody: message.getContent()
            }
        );
      }
      props.setProperty('lastAuthEmailDate', today);
    }
  } else {
    // Authorization has been granted, so continue to respond to the trigger.
    // Main trigger logic here.
    
    // Todo ok, seguimos
    // Desencadenamos acciones en función de las preferencias guardadas
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var filaModelo = +PropertiesService.getDocumentProperties().getProperty('fila');
    var filaRespuesta = e.range.getRow();   
    
    // ¿Aplicar tratamiento a todas las respuestas o solo la última?
    var reprocesar = JSON.parse(PropertiesService.getDocumentProperties().getProperty('reprocesar'));
   
    // ¿Aplicar formato?
    if (props.getProperty('autoFormato') == 'true') {extenderFormato(filaModelo, filaRespuesta, reprocesar);}
    
    // ¿Aplicar fórmulas?
    if (props.getProperty('autoFormula') == 'true') {extenderFormulas(filaModelo, filaRespuesta, reprocesar);}
    
    // ¿Aplicar validación?
    if (props.getProperty('autoValidacion') == 'true') {extenderValidacion(filaModelo, filaRespuesta, reprocesar);}
   
    // ¿Última respuesta a la primera posición?
    if (props.getProperty('autoInversion') == 'true') {
   
      // mover respuesta a primera posición
      if ((lastRow > filaModelo) && (filaRespuesta == lastRow)) {
      
        // Solo se mueve la fila si hay más de 1 respuesta y
        // no se trata de una edición de respuesta previa vía formulario
        var rango = sheet.getRange("A" + lastRow.toString() + ":" + lastRow.toString());
        sheet.moveRows(rango, filaModelo);
      }
    }
  }
}