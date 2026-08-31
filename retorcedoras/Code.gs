//====================================================
// REGISTRO RPM - RETORCEDORAS
// FIBRAS INDUSTRIALES S.A.
//====================================================

//----------------------------------
// CONSTANTES
//----------------------------------
const NOMBRE_HOJA_REGISTROS = "REGISTRO DE MEDICIONES";
const NOMBRE_HOJA_REFERENCIA = "RPM DE REFERENCIA";
const NOMBRE_HOJA_OPERADORES = "OPERADORES";

const POSICION_MINIMA = 1;
const POSICION_MAXIMA = 40;

const OPERADORES = {
  "00486": "UNTIVEROS RAMIREZ EDGAR",
  "00406": "ABAD ARROYO FRANK GERSON",
  "00809": "CHAVARRIA LUQUE SEBASTIAN HERNAN",
  "00539": "RODRIGUEZ VALLE UBER",
  "00681": "SANTIAGO VARGAS BERVELY ARIANA",
  "00183": "SANTILLAN NUÑEZ CLAUDIO",
};

//----------------------------------
// CARGAR FORMULARIO
//----------------------------------
/**
 * Punto de entrada del Web App: sirve el formulario HTML.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index").setTitle(
    "REGISTRO RPM - RETORCEDORAS",
  );
}

//----------------------------------
// VALIDAR OPERADOR
//----------------------------------
/**
 * Valida un código de operador contra la hoja "OPERADORES".
 * @param {string} codigo Código ingresado por el usuario.
 * @return {{valido: boolean, nombre: string}}
 */
function validarCodigo(codigo) {
  codigo = String(codigo).trim();

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_OPERADORES,
  );

  if (!hoja) return { valido: false, nombre: "" };

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    // Starts at 1 for skipping the headings
    const [codigoFila, apellidos, nombres] = datos[i];

    if (String(codigoFila).trim() === codigo) {
      const nombreCompleto =
        String(apellidos || "").trim() + ", " + String(nombres || "").trim();

      return { valido: true, nombre: nombreCompleto };
    }
  }

  return { valido: false, nombre: "" };
}

//----------------------------------
// FECHA - HORA - TURNO
//----------------------------------
/**
 * Determina la fecha, hora y turno actuales (zona horaria Lima).
 * Turno 1: 06:30 - 14:29 | Turno 2: 14:30 - 22:29 | Turno 3: resto.
 * @return {{fecha: string, hora: string, turno: string}}
 */
function obtenerFechaHora() {
  const ZONA_HORARIA = "America/Lima";
  const ahora = new Date();

  const fecha = Utilities.formatDate(ahora, ZONA_HORARIA, "dd/MM/yyyy");
  const hora = Utilities.formatDate(ahora, ZONA_HORARIA, "HH:mm:ss");

  const horas = Number(Utilities.formatDate(ahora, ZONA_HORARIA, "HH"));
  const minutos = Number(Utilities.formatDate(ahora, ZONA_HORARIA, "mm"));
  const horaDecimal = horas + minutos / 60;

  let turno;
  if (horaDecimal >= 6.5 && horaDecimal < 14.5) {
    turno = "TURNO 1";
  } else if (horaDecimal >= 14.5 && horaDecimal < 22.5) {
    turno = "TURNO 2";
  } else {
    turno = "TURNO 3";
  }

  return { fecha, hora, turno };
}

//----------------------------------
// SEPARAR MÁQUINAS AGRUPADAS
//----------------------------------
/**
 * Expande un nombre de máquina agrupado en sus máquinas individuales.
 * Ej: "GALAN 1 Y 2" -> ["GALAN 1", "GALAN 2"]
 * Ej: "DONGTAI 6"    -> ["DONGTAI 6"] (sin cambios)
 * @param {string} texto Nombre de máquina tal como aparece en la hoja.
 * @return {string[]}
 */
function separarMaquinasAgrupadas(texto) {
  // Prefijo de letras (y espacios) seguido de una lista de
  // números separados por comas y/o la palabra "Y"
  const PATRON_GRUPO = /^([A-ZÀ-Ú\s]+?)\s+((?:\d+\s*(?:,\s*|Y\s+))*\d+)$/;
  const coincide = texto.match(PATRON_GRUPO);

  if (!coincide) return [texto];

  const prefijo = coincide[1].trim();
  const numeros = coincide[2].match(/\d+/g);

  return numeros.map((numero) => prefijo + " " + numero);
}

//----------------------------------
// CARGAR TABLA DE RPM DE REFERENCIA
//----------------------------------
/**
 * Lee la hoja "RPM DE REFERENCIA" y devuelve una lista plana de
 * combinaciones máquina + material + título + RPM configurado.
 * Reconstruye los valores de celdas combinadas (máquina y material)
 * y expande las filas que agrupan varias máquinas.
 * @return {Array<{maquina: string, material: string, titulo: string, rpm: number}>}
 */
function cargarTablaRPM() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_REFERENCIA,
  );
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const filas = [];

  let maquinasActuales = [];
  let materialActual = "";

  for (let i = 1; i < datos.length; i++) {
    // Empieza en 1 para saltar el encabezado
    const [maquina, material, titulo, rpm] = datos[i];

    // Las columnas MÁQUINA y MATERIAL están combinadas (merged cells),
    // por eso se "arrastra" el último valor no vacío hacia abajo.
    if (maquina !== "" && maquina !== null) {
      maquinasActuales = separarMaquinasAgrupadas(
        String(maquina).trim().toUpperCase(),
      );
    }
    if (material !== "" && material !== null) {
      materialActual = String(material).trim().toUpperCase();
    }

    const filaIncompleta =
      titulo === "" ||
      titulo === null ||
      rpm === "" ||
      rpm === null ||
      isNaN(Number(titulo)) ||
      isNaN(Number(rpm));

    // Se ignoran filas "PENDIENTE" o sin título/RPM numérico
    if (filaIncompleta) continue;

    // Si la fila agrupa varias máquinas (ej. "GALAN 1 Y 2"),
    // se registra el mismo dato para cada máquina por separado.
    maquinasActuales.forEach((maquinaActual) => {
      filas.push({
        maquina: maquinaActual,
        material: materialActual,
        titulo: String(Number(titulo)),
        rpm: Number(rpm),
      });
    });
  }

  return filas;
}

//----------------------------------
// BUSCAR RPM CONFIGURADO POR DEFECTO
//----------------------------------
/**
 * Busca el RPM configurado por defecto para una combinación exacta
 * de máquina + material + título.
 * @param {string} maquina
 * @param {string} material
 * @param {string} titulo
 * @return {{encontrado: boolean, rpm?: number}}
 */
function obtenerRPM(maquina, material, titulo) {
  maquina = String(maquina).trim().toUpperCase();
  material = String(material).trim().toUpperCase();
  titulo = String(Number(titulo));

  const fila = cargarTablaRPM().find(
    (f) =>
      f.maquina === maquina && f.material === material && f.titulo === titulo,
  );

  return fila ? { encontrado: true, rpm: fila.rpm } : { encontrado: false };
}

//----------------------------------
// OBTENER TÍTULOS DISPONIBLES
//----------------------------------
/**
 * Devuelve los títulos disponibles (ordenados) para una máquina y
 * material dados, según lo registrado en "RPM DE REFERENCIA".
 * @param {string} maquina
 * @param {string} material
 * @return {string[]}
 */
function obtenerTitulos(maquina, material) {
  maquina = String(maquina).trim().toUpperCase();
  material = String(material).trim().toUpperCase();

  const titulos = cargarTablaRPM()
    .filter((f) => f.maquina === maquina && f.material === material)
    .map((f) => Number(f.titulo));

  const titulosUnicos = Array.from(new Set(titulos)).sort((a, b) => a - b);

  return titulosUnicos.map(String);
}

//----------------------------------
// PROMEDIO HISTÓRICO DE RPM REAL
//----------------------------------
/**
 * Calcula el promedio de RPM Real (solo mediciones de HUSO) registrado
 * en "REGISTRO DE MEDICIONES" para una combinación exacta de
 * máquina + material + título, y lo redondea al múltiplo de 50 más cercano
 * @param {string} maquina
 * @param {string} material
 * @param {string} titulo
 * @return {{encontrado: boolean, promedio?: number, cantidadRegistros?: number}}
 */
function calcularPromedioRPMReal(maquina, material, titulo) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_REGISTROS,
  );
  if (!hoja || hoja.getLastRow() < 2) return { encontrado: false };

  maquina = String(maquina).trim().toUpperCase();
  material = String(material).trim().toUpperCase();
  titulo = String(Number(titulo));

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  const idx = {
    maquina: encabezados.indexOf("Máquina"),
    material: encabezados.indexOf("Material"),
    titulo: encabezados.indexOf("Título"),
    tipo: encabezados.indexOf("Tipo"),
    rpmReal: encabezados.indexOf("RPM Real"),
  };

  // Si la hoja no tiene el formato esperado, no se puede calcular.
  if (Object.keys(idx).some((clave) => idx[clave] === -1)) {
    return { encontrado: false };
  }

  let suma = 0;
  let conteo = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const filaMaquina = String(fila[idx.maquina]).trim().toUpperCase();
    const filaMaterial = String(fila[idx.material]).trim().toUpperCase();
    const filaTitulo = String(Number(fila[idx.titulo]));
    const filaTipo = String(fila[idx.tipo]).trim().toUpperCase();

    if (
      filaMaquina !== maquina ||
      filaMaterial !== material ||
      filaTitulo !== titulo ||
      filaTipo !== "HUSO"
    ) {
      continue;
    }

    const rpmReal = Number(fila[idx.rpmReal]);
    if (!isNaN(rpmReal) && fila[idx.rpmReal] !== "") {
      suma += rpmReal;
      conteo++;
    }
  }

  if (conteo === 0) return { encontrado: false };

  const promedio = suma / conteo;
  const promedioRedondeado = Math.round(promedio / 50) * 50;

  return {
    encontrado: true,
    promedio: promedioRedondeado,
    cantidadRegistros: conteo,
  };
}

/**
 * Punto de entrada para el formulario: obtiene el promedio histórico
 * de RPM Real para la combinación indicada. Se usa tanto para mostrar
 * la referencia en pantalla como, para DONGTAI 4 y 5, para calcular
 * el rango aceptable real.
 * @param {string} maquina
 * @param {string} material
 * @param {string} titulo
 * @return {{encontrado: boolean, promedio?: number, cantidadRegistros?: number}}
 */
function obtenerRPMPromedio(maquina, material, titulo) {
  return calcularPromedioRPMReal(maquina, material, titulo);
}

//----------------------------------
// VALIDAR REGISTRO
//----------------------------------
/**
 * Valida, en el servidor, que un registro individual tenga una
 * posición numérica dentro del rango permitido. Es una segunda
 * capa de validación además de la del formulario.
 * @param {{posicion: string|number}} registro
 * @return {boolean}
 */
function posicionEsValida(registro) {
  const posicion = Number(registro.posicion);
  return (
    Number.isInteger(posicion) &&
    posicion >= POSICION_MINIMA &&
    posicion <= POSICION_MAXIMA
  );
}

/**
 * Calcula el estado (dentro/fuera de rango) que se guardará para una
 * medición. Solo aplica a HUSO; CORDEL se registra sin evaluar.
 * El estado ya fue calculado en el cliente (usando el RPM Configurado
 * o el RPM Promedio, según la máquina) al momento de ingresar el RPM
 * Real, por lo que aquí solo se traduce el valor recibido; no se
 * recalcula para evitar trabajo redundante en el servidor.
 * @param {{tipo: string, estado: string}} registro
 * @return {string}
 */
function calcularEstadoFinal(registro) {
  const esHuso = String(registro.tipo).toUpperCase() === "HUSO";
  if (!esHuso) return "";

  const fueraDeRango = String(registro.estado).toUpperCase().includes("FUERA");
  return fueraDeRango ? "FUERA DE RANGO" : "DENTRO DE RANGO";
}

/**
 * Calcula la observación final a guardar. Para HUSO, si viene vacía
 * se completa con "OK"; para CORDEL se respeta el valor tal cual
 * (puede quedar en blanco).
 * @param {{tipo: string, observacion: string}} registro
 * @return {string}
 */
function calcularObservacionFinal(registro) {
  const esHuso = String(registro.tipo).toUpperCase() === "HUSO";
  const observacion = String(registro.observacion || "").trim();

  if (observacion === "" && esHuso) return "OK";
  return observacion.toUpperCase();
}

//----------------------------------
// GUARDAR REGISTRO
//----------------------------------
/**
 * Guarda un lote de mediciones (8 por turno) en la hoja de registros.
 * Crea los encabezados si la hoja está vacía.
 * @param {object} datos Payload enviado desde el formulario.
 * @return {{ok: boolean, mensaje: string}}
 */
function guardarRegistro(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_REGISTROS,
  );

  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      "Fecha",
      "Hora",
      "Turno",
      "Código",
      "Operador",
      "Máquina",
      "Material",
      "Título",
      "RPM Configurada",
      "RPM Promedio",
      "Lado",
      "Tipo",
      "Posición",
      "RPM Real",
      "Estado RPM",
      "Observación",
    ]);
  }

  const registrosInvalidos = datos.registros.filter(
    (registro) => !posicionEsValida(registro),
  );
  if (registrosInvalidos.length > 0) {
    return {
      ok: false,
      mensaje:
        "La posición debe ser un número entero entre " +
        POSICION_MINIMA +
        " y " +
        POSICION_MAXIMA +
        ".",
    };
  }

  const info = obtenerFechaHora();

  // El RPM Promedio ya fue calculado y mostrado en el formulario al
  // seleccionar máquina + material + título; se guarda tal cual lo
  // envía el cliente para no recalcularlo (y potencialmente obtener
  // un valor distinto) en este momento.
  const rpmPromedioParaGuardar =
    datos.rpmPromedio !== undefined &&
    datos.rpmPromedio !== null &&
    String(datos.rpmPromedio).trim() !== ""
      ? Number(datos.rpmPromedio)
      : "";

  const filas = datos.registros.map((registro) => [
    info.fecha,
    info.hora,
    info.turno,
    datos.codigo,
    datos.nombre,
    datos.maquina,
    datos.material,
    datos.titulo,
    datos.rpmConfigurado,
    rpmPromedioParaGuardar,
    registro.lado,
    String(registro.tipo).toUpperCase(),
    registro.posicion,
    Number(registro.rpm),
    calcularEstadoFinal(registro),
    calcularObservacionFinal(registro),
  ]);

  hoja
    .getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
    .setValues(filas);

  return { ok: true, mensaje: "Registro guardado correctamente." };
}
