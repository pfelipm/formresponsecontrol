// Ámbito de autorización
/**
 * @OnlyCurrentDoc
 */
 
function onInstall(e) {
  
  // Otras cosas que se deben hacer siempre
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('✅ Configurar', 'configurar')
    .addItem('❔ Comprobar estado', 'comprobarEstado') 
    .addToUi();
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

function desactivar() {
  
  // Comprueba si ya hay un trigger ON_FORM_SUBMIT asociado al proyecto ¡para el usuario actual!
  var estado;
  var triggerDe = PropertiesService.getDocumentProperties().getProperty('triggerDe');
  if (triggerDe == '' || triggerDe == null) {
    SpreadsheetApp.getUi().alert('💡 Form Response Control no está activado.\n\n¡Nada que hacer!');
    estado = false;
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
      ScriptApp.deleteTrigger(triggerActual);
      // Actualizar registro de propietario del trigger
      // Se asume que la lógica del código impide activaciones múltiples
      // Por parte de diferentes usuarios
      PropertiesService.getDocumentProperties().setProperty('triggerDe','');
      SpreadsheetApp.getUi().alert('🛑 Form Response Control ha sido desactivado.');
      estado = false;
    }
    else {
      // Solo queda un caso, otro usuario ha activado el trigger, no podemos hacer nada
      SpreadsheetApp.getUi().alert('💡 Form Response Control ha sido activado por:\n\n' + triggerDe + '\n\n¡Pídele que lo desactive!');
      estado = true;
    }
  }
  return estado;
}

function reActivar() {
  
  // Comprobar si otro editor de la hdc ya ha instalado el trigger
  var estado;
  var triggerDe = PropertiesService.getDocumentProperties().getProperty('triggerDe');
  if (triggerDe != '' && triggerDe != null ) {
    SpreadsheetApp.getUi().alert('💡 Form Response Control ya ha sido activado por:\n\n' + triggerDe + '\n\n¡Nada que hacer!');
    estado = true;
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
      PropertiesService.getDocumentProperties().setProperty('triggerDe',Session.getEffectiveUser())
      SpreadsheetApp.getUi().alert('🚀 Form Response Control ha sido activado.');
      estado = true;
    }
    catch (e) {SpreadsheetApp.getUi().alert('¡Error!','Se han producido errores activando el complemento, es posible que no funcione correctamente:\n\n'+e,SpreadsheetApp.getUi().ButtonSet.OK);}
  }
  return estado;
}

function configurar() {

  // Si es la 1ª vez que se ejecuta, inicializar ajustes (propiedades del documento)
  if (PropertiesService.getDocumentProperties().getProperty('triggerDe') == null) {
    PropertiesService.getDocumentProperties().setProperties({
      'autoFormato' : 'false',
      'autoFormula' : 'false',
      'autoInversion' : 'false',
      'triggerDe' : '',
    }, true);
  }
  
  // Script ya configurado, abrimos el panel de configuración
  var panel=HtmlService.createHtmlOutputFromFile('panelLateral')
    .setTitle('✅ Configuración FRC');
  SpreadsheetApp.getUi().showSidebar(panel);  
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
    // SpreadsheetApp.getUi().alert(preferencias[ajuste].clave + ' ' + propiedadesDocumento.getProperty(preferencias[ajuste].clave));
  }
  // SpreadsheetApp.getUi().alert('Los ajustes se han guardado.');
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
   
    if (props.getProperty('autoFormato') == 'true') {
   
      // extender formato de primera fila al resto de respuestas  
      if (lastRow > 2) {
      
        // Formato
        sheet.getRange(2, 1, 1, lastColumn).copyFormatToRange(sheet, 1, lastColumn, 3, lastRow);
        
        // Altura de fila
        var alturaFila = sheet.getRowHeight(2);
        sheet.setRowHeights(3, lastRow - 2, alturaFila);     
      }   
    }
   
    if (props.getProperty('autoFormula') == 'true') {
   
      // copiar fórmulas de primera fila
      if (lastRow > 2) {
        for (var col = 1; col <= lastColumn; col++) {
          celdaFormula = sheet.getRange(2,col,1,1);
          
          // Si en alguna celda de la fila 2 hay una fórmula la copiamos a la última
          if (celdaFormula.getFormula() != '') {celdaFormula.copyTo(sheet.getRange(lastRow,col));}
        }
      }     
    }
   
    if (props.getProperty('autoInversion') == 'true') {
   
      // mover respuesta a primera posición
      if (lastRow > 2) {
        var rango = sheet.getRange("A" + lastRow.toString() + ":" + lastRow.toString());
        sheet.moveRows(rango, 2);
      }
    }
  }
}