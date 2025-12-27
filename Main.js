async function main(workbook: ExcelScript.Workbook) {
  reg2log(workbook);

  // Genera los arrays-JSON para cada tabla
  // De la hoja Interfaz
  let promos = ap2Json(workbook, "E14:J23");
  reg2log(workbook, "E", JSON.stringify(promos));

  let capitulos = cap2Json(workbook, "E35:K44");
  reg2log(workbook, "F", JSON.stringify(capitulos));

  let programas = pgm2Json(workbook, "E56:G65");
  reg2log(workbook, "G", JSON.stringify(programas));

  // De la hoja Home
  let variables = env2Json(workbook, "B6:C9");
  reg2log(workbook, "C", JSON.stringify(variables));

  let correos = directorio2Json(workbook, "F6:H24");
  reg2log(workbook, "D", JSON.stringify(correos));

  // Contruir el esquema del JSON
  let resultadoFinal = {
    data: variables,
    usuarios: correos,
    autopromos: promos,
    capitulos: capitulos,
    programas: programas,
  };


  ////////////////////  ENVIO DE PETICIÓN HTTP

  const flowUrl = "https://default872ec208871247cc805226a93f2e1f.5b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/ed88d3085bd24caea2cc7f0d7ef7b4a4/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2MNHlpduWj0NF5VLiuLzwViHEzVfI46gkTVBB2gJJvI"; // Opciones de petición fetch

  const requestOptions: RequestInit = {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(resultadoFinal),
  }; // Enviar HTTP POST al flujo Power Automate

  try {

    const response = await fetch(flowUrl, requestOptions);

    if (!response.ok) {
      throw new Error(
        `Error en el flujo: ${response.status} - ${await response.text()}`

      );
    }

    // Registro de Log de finalización
    reg2log(workbook, "B", "Ejecutado correctamente.");

    Copy2DB(workbook, "E14:E23", "D14:J14", "Autopromos");    // Copia Autopromos
    Copy2DB(workbook, "E35:E44", "C35:K35", "Capitulos");    // Copia Capitulos
    Copy2DB(workbook, "E56:E65", "D56:G56", "Programas");    // Copia Capitulos

  //LimpiarInterfaz(workbook);

  } catch (e) {
    // Manejo de errores, opcional guardar el mensaje de error en alguna celda o log
    console.log("Error en envío al flujo: " + e.message);
    reg2log(workbook, "B", e.message + "    |||   " + JSON.stringify(resultadoFinal));
  }

  ////////////////////  FIN DE ENVIO DE PETICIÓN HTTP



}

//Función para obtener las variables de entorno
function env2Json(workbook: ExcelScript.Workbook, rango: string): { [key: string]: string } {
  let sheet = workbook.getWorksheet("Home");
  let values = sheet.getRange(rango).getValues();
  let result: { [key: string]: string } = {};

  // 1. Procesar los valores del rango (clave-valor)
  for (let row of values) {
    let key = row[0];
    let value = row[1];

    if (!key || String(key).trim() === "") {
      continue; // Ignorar filas sin clave
    }
    result[String(key)] = value ? String(value) : "";
  }

  // --- NUEVAS VARIABLES A AGREGAR AL RESULTADO ---

  // Obtener y preparar el nombre del archivo
  const fileName = workbook.getName();
  // Elimina la extensión (Ej: "Libro.xlsx" -> "Libro")
  const fileNameSinExtension = fileName.replace(/\.[^/.]+$/, "");

  // Obtener propiedades del libro
  const propiedades = workbook.getProperties();
  const UltimoEditor = propiedades.getLastAuthor();

  // Obtener la fecha y hora actual
  const now = new Date();
  // Convertir la fecha a una cadena para almacenarla en el objeto { [key: string]: string }
  const fechaActual = now.toISOString(); // Formato estándar ISO para fácil procesamiento

  // 2. Agregar las nuevas variables al objeto result
  // Se agregan como nuevas entradas de clave-valor

  result["FileName"] = fileName; // Nombre completo del archivo
  result["FileNameWithoutExtension"] = fileNameSinExtension;
  result["LastAuthor"] = UltimoEditor;
  result["CurrentDateTimeUTC"] = fechaActual; // Usando el formato ISO

  // 3. Devolver el objeto final
  return result;
}

//Función para obtener los correos de destino y CC
function directorio2Json(workbook: ExcelScript.Workbook, rango: string) {
  let sheet = workbook.getWorksheet("Home");
  let values = sheet.getRange(rango).getValues();

  let destinatarios: string[] = [];
  let copiados: string[] = [];
  let sbmails: string[] = [];

  for (let row of values) {
    let to = row[0];
    let cc = row[1];
    let sb = row[2];

    // Validar y agregar solo si hay dato y parece un email
    if (to && String(to).includes("@")) {
      destinatarios.push(String(to));
    }
    if (cc && String(cc).includes("@")) {
      copiados.push(String(cc));
    }
    if (sb && String(sb).includes("@")) {
      sbmails.push(String(sb));
    }
  }
  const destinatariosString = destinatarios.join(';');
  const copiadosString = copiados.join(';');
  const sbmailsString = sbmails.join(';');

  return { destinatariosString, copiadosString, sbmailsString };
}

// Función para la primera tabla ("AUTOPROMOS")
function ap2Json(workbook: ExcelScript.Workbook, rango: string) {
  let worksheet = workbook.getWorksheet("Interfaz") //getActiveWorksheet();
  const range = worksheet.getRange(rango).getValues();
  return range.filter(row => row.some(cell => cell !== "" && cell !== null)).map((row) => ({
    CÓDIGO_ASIGNADO: row[0],
    PRODUCTO: row[1],
    REFERENCIA: row[2],
    FECHA_DE_VIGENCIA: typeof row[3] === "number" ? excelSerial2Date(row[3]) : row[3],
    DURACIÓN: row[4],
    FECHA_EMISIÓN: typeof row[5] === "number" ? excelSerial2Date(row[5]) : row[5],
  }));
}

// Función para la segunda tabla ("CAPÍTULOS")
function cap2Json(workbook: ExcelScript.Workbook, rango: string) {
  let worksheet = workbook.getWorksheet("Interfaz") //getActiveWorksheet();
  const range = worksheet.getRange(rango).getValues();
  return range.filter(row => row.some(cell => cell !== "" && cell !== null)).map((row) => ({
    CÓDIGO_ASIGNADO: row[0],
    PROGRAMA: row[1],
    REFERENCIA: row[2],
    DURACIÓN: row[3],
    FECHA_VIGENCIA: typeof row[4] === "number" ? excelSerial2Date(row[4]) : row[4]
  }));
}

//Función para la tercera tabla("PROGRAMAS")
function pgm2Json(workbook: ExcelScript.Workbook, rango: string) {
  let worksheet = workbook.getWorksheet("Interfaz") //getActiveWorksheet();
  const range = worksheet.getRange(rango).getValues();
  return range.filter(row => row.some(cell => cell !== "" && cell !== null)).map((row) => ({
    CÓDIGO: row[0],
    PROGRAMA: row[1],
    CORTINILLAS: row[2]
  }));

}

// Funcion registrar info en log dependiendo de la columna ingresada.  
function reg2log(workbook: ExcelScript.Workbook, idcolumna: string = "A", textodatos: string = "") {
  const hoja = workbook.getWorksheet("Log") ?? workbook.addWorksheet("Log");
  const ultimaA = hoja.getRange("A:A").getUsedRange();
  let ultimaFila = ultimaA ? ultimaA.getRowCount() + (idcolumna === "A" ? 1 : 0) : 1;

  if (idcolumna === "A") {
    const ahora = new Date();
    textodatos = `${String(ahora.getDate()).padStart(2, '0')}/${String(ahora.getMonth() + 1).padStart(2, '0')}/${String(ahora.getFullYear()).slice(-2)} | ${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}:${String(ahora.getSeconds()).padStart(2, '0')}`;
  }

  hoja.getRange(`${idcolumna}${ultimaFila}`).setValue(textodatos);
}

// Funcion para almacenar la info de las tablas en la DB de cada material.
function Copy2DB(workbook: ExcelScript.Workbook, idrange: string, fromcopy: string, destiny: string) {
  const interfaz = workbook.getWorksheet("Interfaz");

  const fromCopy = interfaz.getRange(fromcopy);  // Define el tama;o de lo que se va a copiar.
  const lengthCopy = fromCopy.getColumnCount();
  const postCopy = fromCopy.getColumnIndex();

  const idRange = interfaz.getRange(idrange);
  const valueIdRange = idRange.getValues();
  const indexRow = idRange.getRowIndex() + 1;

  const sDestiny = workbook.getWorksheet(destiny) ?? workbook.addWorksheet(destiny);
  let next2Use = sDestiny.getRange("A:A").getUsedRange().getRowCount();
  let count = 0;

  valueIdRange.forEach((fila, i) => {
    const v = fila[0];
    const estaVacia = v === null || v === undefined || String(v).trim() === "";
    if (!estaVacia) {
      let src = interfaz.getRangeByIndexes(indexRow + i - 1, postCopy, 1, lengthCopy);
      let dst = sDestiny.getRangeByIndexes(next2Use, 1, 1, lengthCopy);

      // Copiar valores (y si quieres formatos, agrega otra copia de formatos)
      dst.copyFrom(src, ExcelScript.RangeCopyType.values, false, false);

      next2Use += 1; // avanzamos la fila destino para la próxima copia\
      count += 1;
    }
  });

  // Agregar ID a la info copiada.
  for (let i = 1; i <= count; i++) {
    sDestiny.getRange(`A${next2Use - count + i}`).setValue(next2Use - count + i - 1);
  }
}


function LimpiarInterfaz(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getWorksheet("Interfaz");
  selectedSheet.getRange("E14:I23").clear(ExcelScript.ClearApplyTo.contents);	// Autopromos
  selectedSheet.getRange("P14:U23").clear(ExcelScript.ClearApplyTo.contents);	// Capitulos
  selectedSheet.getRange("E37:E46").clear(ExcelScript.ClearApplyTo.contents);	// Programa
}


// FUNCIONES EXTRA

//Convertir el numero derial que arroja excel en fechas
function excelSerial2Date(serial: number): string {
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  const day = String(date.getDate() + 1).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0'); // Mes empieza en 0
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}
