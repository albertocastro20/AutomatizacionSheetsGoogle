//nombre del archivo:  Código.gs

/**
 * @fileoverview Este script automatiza la gestión de tareas en Google Workspace
 * integrando Google Sheets, Gmail y Google Calendar.
 * Se activa al editar una hoja de cálculo para enviar notificaciones
 * y crear eventos de calendario.
 * Castro López Cristian Alberto
 */

// --- Configuración Global ---
const NOMBRE_HOJA = 'Tareas'; // Nombre de la hoja donde están los datos 
const FILA_ENCABEZADOS = 1;         // Fila donde se encuentran los encabezados de las columnas
const NOMBRE_PROYECTO_COL = 1; // Columna A
const TAREA_COLUMNA = 2;          // Columna B
const CORREO_COL = 3;    // Columna E (Correo electrónico del asignado)
const STATUS_COL = 4;        // Columna D
const FECHA_COL = 5;      // Columna C
const EVENTO_CALENDARIO_COL = 6; // Columna F (Casilla de verificación para evento de calendario)
const CORREO_ENVIADO_COL = 7;    // Columna G: Nueva casilla de verificación para "Correo Enviado"


// --- Funciones Principales ---

/**
 * Función principal que se activa automáticamente al editar la hoja de cálculo.
 * Esta función es el punto de entrada para los activadores 'On edit'.
 *
 * El parametro e funge como @param {GoogleAppsScript.Events.SheetsOnEdit}
 */
function onEditTrigger(e) {
  Logger.log('Activador edicionHoja ejecutado.');

  //Evalua el evento
  if (!e || !e.range) {
    Logger.log('Error: Objeto de evento o rango no válido.');
    return;
  }
//Almacenamos el valor del rango y de la hoja
  const rango = e.range;
  const hoja = rango.getSheet();

//Evaluamos si la hoja es la correcta
  if (hoja.getName() !== NOMBRE_HOJA) {
    Logger.log(`Edición fuera de la hoja principal: ${hoja.getName()}. Ignorando.`);
    return;
  }
//Evaluamos si la edicion no es en la fila de encabezados
  if (rango.getRow() === FILA_ENCABEZADOS) {
    Logger.log('Edición en la fila de encabezados. Ignorando.');
    return;
  }

//Almacenamos la fila que está siendo editada
  const filaEditada = rango.getRow();
  // const editedColumn = range.getColumn(); // No la usaremos para la lógica principal

  Logger.log(`Edición detectada en la fila: ${filaEditada}`);

  // Obtener todos los datos de la fila editada.
  // Es importante obtener TODAS las columnas que necesitamos, incluyendo las de control.
  const filaDatos = hoja.getRange(filaEditada, 1, 1, hoja.getLastColumn()).getValues()[0];

  const nombreProyecto = filaDatos[NOMBRE_PROYECTO_COL - 1];
  const tarea = filaDatos[TAREA_COLUMNA - 1];
  const fecha = filaDatos[FECHA_COL - 1];
  const status = filaDatos[STATUS_COL - 1];
  const correo = filaDatos[CORREO_COL - 1];
  // Convertimos el valor de la casilla de verificación a booleano de forma segura
  const calendarioEvento = Boolean(filaDatos[EVENTO_CALENDARIO_COL - 1]);
  const correoEvento = Boolean(filaDatos[CORREO_ENVIADO_COL - 1]); 

  Logger.log(`Datos de la fila ${filaEditada}: Proyecto: ${nombreProyecto}, Tarea: ${tarea}, Estado: ${status}, Fecha: ${fecha}, Asignado: ${correo}, Evento Creado: ${calendarioEvento}, Correo Enviado: ${correoEvento}`);


  // --- Lógica de Automatización (Disparo al completar campos relevantes) ---

  // 1. Automatización de Notificaciones (Gmail)
  // Requisitos: Llenado de las columnas: Proyecto, Tarea, Estado, Correo Asignado y que el correo NO se haya enviado ya.

  //Evaluamos si las columnas están llenas para procedes 
  if (nombreProyecto && tarea && status && correo && correoValido(correo)) {
    const statusMinusculas = status.toString().toLowerCase(); //convertimos a minuscolas

    // enviar correo de Tarea pendiente - Solo una vez
    // Se envía si el estado es 'pendiente' y NO se ha enviado correo para esta fila
    if (statusMinusculas === 'pendiente' && !correoEvento) {
      Logger.log(`Tarea '${tarea}' en proyecto '${nombreProyecto}' marcada como 'Pendiente'. Enviando notificación inicial.`);
      enviarPendienteEmail(nombreProyecto, tarea, correo);
      // Marcar la casilla "Correo Enviado" para que no se reenvíe este mismo correo.
      hoja.getRange(filaEditada, CORREO_ENVIADO_COL).setValue(true);
      Logger.log(`Casilla 'Correo Enviado' marcada en la fila ${filaEditada}.`);
    }
    
    // enviar correo de Tarea COMPLETADA - Cada vez que cambie a 'completado'

    else if (statusMinusculas === 'completado') {
      Logger.log(`Tarea '${tarea}' en proyecto '${nombreProyecto}' marcada como 'Completado'.`);
      enviarCompletadoEmail(nombreProyecto, tarea, correo);

    }
  } else {
    Logger.log(`No se puede enviar correo de notificación. Faltan datos (Proyecto, Tarea, Estado, Email) o email inválido para la fila ${filaEditada}.`);
  }


  // 2. Integración con Google Calendar
  // Requisitos: Columnas llenas:  Proyecto, Tarea, Fecha de Vencimiento y que el evento NO se haya creado ya.
  if (nombreProyecto && tarea && fecha instanceof Date && !isNaN(fecha.getTime()) && !calendarioEvento) {
    Logger.log(`Fecha de vencimiento para '${tarea}' detectada y evento no creado. Creando evento de calendario.`);
    crearEventoCalendario(nombreProyecto, tarea, fecha, filaEditada, hoja);
  }
  //Se verifica si ya existe la instancia con la casilla de verificación
   else if (nombreProyecto && tarea && fecha instanceof Date && !isNaN(fecha.getTime()) && calendarioEvento) {
    Logger.log(`Fecha de vencimiento para '${tarea}' actualizada, pero evento ya creado. Considerar actualizar evento existente.`);
  } else {
    Logger.log(`No se puede crear evento de calendario. Faltan datos (Proyecto, Tarea, Fecha) o evento ya creado para la fila ${filaEditada}.`);
  }
}

/**
 * Envía un correo electrónico de notificación cuando una tarea se marca como completada.
 *
 * @param {string} nombreProyecto El nombre del proyecto.
 * @param {string} tarea La descripción de la tarea.
 * @param {string} correoRecibido La dirección de correo electrónico del destinatario.
 */
function enviarCompletadoEmail(nombreProyecto, tarea, correoRecibido) {
  
  //Se verifica si el correo es válido
  if (!correoRecibido || !correoValido(correoRecibido)) {
    Logger.log(`Advertencia: No se pudo enviar correo. Correo electrónico inválido o vacío para '${tarea}': ${correoRecibido}`);
    return;
  }

//Almacenamos lo que vamos a mandar por correo
  const titulo = `✅ Tarea Completada: ${tarea} (${nombreProyecto})`;
  const texto = `Hola,\n\nLa tarea "${tarea}" del proyecto "${nombreProyecto}" ha sido marcada como COMPLETA.\n\n¡Buen trabajo!\n\nSaludos,\nTu Sistema de Automatización`;

  //Proceso para enviar el correo
  try {
    GmailApp.sendEmail(correoRecibido, titulo, texto);
    Logger.log(`Correo de finalización enviado a ${correoRecibido} para la tarea '${tarea}'.`);
  } catch (error) {
    Logger.log(`Error al enviar correo para '${tarea}' a ${correoRecibido}: ${error.message}`);
  }
}

function enviarPendienteEmail(nombreProyecto, tarea, correoRecibido) {
  //Se verifica si el correo es válido
  if (!correoRecibido || !correoValido(correoRecibido)) {
    Logger.log(`Advertencia: No se pudo enviar correo. Correo electrónico inválido o vacío para '${tarea}': ${correoRecibido}`);
    return;
  }
//Almacenamos lo que vamos a mandar por correo
  const titulo = `🔔 Tarea Asignada: ${tarea} (${nombreProyecto})`;
  const texto = `Hola,\n\nLa tarea "${tarea}" del proyecto "${nombreProyecto}" ha sido marcada como PENDIENTE.\n\n¡Esfuérzate!\n\nSaludos,\nTu Sistema de Automatización`;

//Proceso para enviar el correo
  try {
    GmailApp.sendEmail(correoRecibido, titulo, texto);
    Logger.log(`Correo de Asignación enviado a ${correoRecibido} para la tarea '${tarea}'.`);
  } catch (error) {
    Logger.log(`Error al enviar correo para '${tarea}' a ${correoRecibido}: ${error.message}`);
  }
}

/**
 * Crea un evento en Google Calendar para una tarea con fecha de vencimiento.
 *
 * @param {string} nombreProyecto El nombre del proyecto.
 * @param {string} tarea La descripción de la tarea.
 * @param {Date} fecha La fecha de vencimiento de la tarea (objeto Date).
 * @param {number} fila La fila de la hoja de cálculo donde se encuentra la tarea.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja El objeto de la hoja de cálculo.
 */
function crearEventoCalendario(nombreProyecto, tarea, fecha, fila, hoja) {
  //Verificamos si la fecha es válida
  if (!(fecha instanceof Date) || isNaN(fecha.getTime())) {
    Logger.log(`Error: Fecha de vencimiento inválida para la tarea '${tarea}': ${fecha}. No se creará el evento.`);
    return;
  }

  const calendario = CalendarApp.getDefaultCalendar();
  const titulo = `Vencimiento: ${tarea} (${nombreProyecto})`;
  const descripcion = `Tarea: ${tarea}\nProyecto: ${nombreProyecto}\n\nRevisa la hoja de cálculo para más detalles.`;

  const diaComienzo = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  const diaFin = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate() + 1);

  try {
    const evento = calendario.createEvent(titulo, diaComienzo, diaFin, {
      description: descripcion
    });
    Logger.log(`Evento de calendario creado para '${tarea}': ${evento.getId()}`);

    hoja.getRange(fila, EVENTO_CALENDARIO_COL).setValue(true);
    Logger.log(`Casilla 'Evento Creado' marcada en la fila ${fila}.`);

  } catch (error) {
    Logger.log(`Error al crear evento de calendario para '${tarea}': ${error.message}`);
  }
}

/**
 * Función auxiliar para validar un formato de correo electrónico básico.
 * @param {string} correo La cadena de correo electrónico a validar.
 * @returns {boolean} Verdadero si el correo parece válido, falso en caso contrario.
 */
function correoValido(correo) {
  return /\S+@\S+\.\S+/.test(correo);
}

// --- Notas importantes para el despliegue y uso ---
// 1. Abre tu Google Sheet y ve a Extensiones > Apps Script.
// 2. Copia y pega este código en el archivo 'Código.gs'.
// 3. Asegúrate de que los nombres de las columnas en tu hoja de cálculo coincidan
//    con las constantes `_COL` definidas al inicio del script.
//    *** ¡IMPORTANTE! Asegúrate de añadir a las colunas 6 y 7 casillas de verificación en tu Google Sheet
// 4. Configura el activador 'onEdit':
//    - En el editor de Apps Script, haz clic en el icono del reloj (Activadores).
//    - Haz clic en 'Añadir activador'.
//    - Elige 'onEditTrigger' para la función.
//    - Selecciona 'Desde la hoja de cálculo' como origen del evento.
//    - Selecciona 'Al editar' como tipo de evento.
//    - Guarda el activador. La primera vez, se te pedirá que autorices el script
//      para acceder a tus datos de Google Sheets, Gmail y Calendar.
// 5. Para probar:
//    - Añade una nueva fila y llena todos los datos (Proyecto, Tarea, Correo de gmail válido, Status="Pendiente" o "Completado", Fecha de entrega válida, Evento Creado (déjalo desmarcado)
//      y Correo Enviado (dejar desmarcado)).
//    - Tan pronto como la fila tenga todos esos datos, el correo "Pendiente" se enviará
//      y la casilla "Correo Enviado" se marcará automáticamente.
//    - Si pones una fecha de vencimiento y la casilla "Evento Creado" está desmarcada,
//      el evento de calendario se creará y la casilla "Evento Creado" se marcará.
//    - Luego, cambia el estado a "Completado" y se enviará otro correo.
