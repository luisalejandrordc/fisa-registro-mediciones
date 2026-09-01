//====================================================
// REGISTRO DE CRUCES - TRENZADORAS
// FIBRAS INDUSTRIALES S.A.
//====================================================

//----------------------------------
// CONSTANTES
//----------------------------------
const NOMBRE_HOJA_REGISTROS = "REGISTRO DE MEDICIONES";
const NOMBRE_HOJA_SUPERVISORES = "SUPERVISORES";
const NOMBRE_HOJA_TRENZADORAS = "TRENZADORAS";
const NOMBRE_HOJA_ORDENES = "ORDENES DE TRABAJO";

const MARGEN_TOLERANCIA_CRUCE = 0.05; // ±5%
const CANTIDAD_MEDICIONES = 3; // 3 trenzadoras, de filas diferentes

//----------------------------------
// CARGAR FORMULARIO
//----------------------------------
/**
 * Punto de entrada del Web App: sirve el formulario HTML.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index").setTitle(
    "REGISTRO CRUCES - TRENZADORAS",
  );
}

//----------------------------------
// VALIDAR SUPERVISOR
//----------------------------------
/**
 * Valida un código de supervisor contra la hoja "SUPERVISORES".
 * @param {string} codigo Código ingresado por el usuario.
 * @return {{valido: boolean, nombre: string}}
 */
function validarCodigo(codigo) {
  codigo = String(codigo).trim();

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_SUPERVISORES,
  );
  if (!hoja) return { valido: false, nombre: "" };

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    // Empieza en 1 para saltar el encabezado
    const [codigoFila, apellidos, nombres] = datos[i];

    if (String(codigoFila).trim() === codigo) {
      const nombreCompleto = (
        String(apellidos || "").trim() +
        ", " +
        String(nombres || "").trim()
      ).trim();

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
// CARGAR TABLA DE ÓRDENES DE TRABAJO
//----------------------------------
/**
 * Lee la hoja "ORDENES DE TRABAJO" y devuelve una lista plana de
 * OT + código de producto + nombre de producto + cruces STD.
 * @return {Array<{ordenTrabajo: string, codigo: string, nombreProducto: string, crucesSTD: number}>}
 */
function cargarTablaOrdenes() {
  const hoja =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_ORDENES);
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const filas = [];

  for (let i = 1; i < datos.length; i++) {
    // Empieza en 1 para saltar el encabezado
    const [ordenTrabajo, codigo, nombreProducto, crucesSTD] = datos[i];

    if (ordenTrabajo === "" || ordenTrabajo === null) continue;

    filas.push({
      ordenTrabajo: String(ordenTrabajo).trim().toUpperCase(),
      codigo: String(codigo || "").trim(),
      nombreProducto: String(nombreProducto || "").trim(),
      crucesSTD: Number(crucesSTD),
    });
  }

  return filas;
}

//----------------------------------
// BUSCAR DATOS DE UNA OT
//----------------------------------
/**
 * Busca el nombre de producto y cruces STD configurados para una
 * orden de trabajo exacta.
 * @param {string} ordenTrabajo
 * @return {{encontrado: boolean, nombreProducto?: string, crucesSTD?: number}}
 */
function obtenerDatosOT(ordenTrabajo) {
  ordenTrabajo = String(ordenTrabajo).trim().toUpperCase();

  const fila = cargarTablaOrdenes().find(
    (f) => f.ordenTrabajo === ordenTrabajo,
  );

  if (!fila || isNaN(fila.crucesSTD)) {
    return { encontrado: false };
  }

  return {
    encontrado: true,
    nombreProducto: fila.nombreProducto,
    crucesSTD: fila.crucesSTD,
  };
}

//----------------------------------
// CARGAR TABLA DE TRENZADORAS
//----------------------------------
/**
 * Lee la hoja "TRENZADORAS" y devuelve una lista plana de
 * fila + número. Reconstruye los valores de la columna FILA, que
 * está combinada (merged cells) y se "arrastra" hacia abajo.
 * @return {Array<{fila: string, numero: string}>}
 */
function cargarTablaTrenzadoras() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    NOMBRE_HOJA_TRENZADORAS,
  );
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const filas = [];

  let filaActual = "";

  for (let i = 1; i < datos.length; i++) {
    // Empieza en 1 para saltar el encabezado
    const [fila, numero] = datos[i];

    // La columna FILA está combinada (merged cells), por eso se
    // "arrastra" el último valor no vacío hacia abajo.
    if (fila !== "" && fila !== null) {
      filaActual = String(fila).trim().toUpperCase();
    }

    if (numero === "" || numero === null) continue;

    // El número no siempre es numérico (ej. "A1"), se guarda como texto
    filas.push({
      fila: filaActual,
      numero: String(numero).trim().toUpperCase(),
    });
  }

  return filas;
}

//----------------------------------
// OBTENER FILAS DISPONIBLES
//----------------------------------
/**
 * Devuelve las filas disponibles (en el orden en que aparecen en la
 * hoja "TRENZADORAS"), sin duplicados.
 * @return {string[]}
 */
function obtenerFilas() {
  const filas = cargarTablaTrenzadoras().map((f) => f.fila);
  return Array.from(new Set(filas));
}

//----------------------------------
// OBTENER N° DE TRENZADORA DISPONIBLES PARA UNA FILA
//----------------------------------
/**
 * Devuelve los números de trenzadora disponibles para una fila dada,
 * en el orden en que aparecen en la hoja "TRENZADORAS".
 * @param {string} fila
 * @return {string[]}
 */
function obtenerNumeros(fila) {
  fila = String(fila).trim().toUpperCase();

  const numeros = cargarTablaTrenzadoras()
    .filter((f) => f.fila === fila)
    .map((f) => f.numero);

  return Array.from(new Set(numeros));
}

//----------------------------------
// CALCULAR ESTADO DEL CRUCE
//----------------------------------
/**
 * Compara el cruce real contra el cruce STD con un margen de
 * tolerancia de ±5%.
 * @param {number|string} cruceReal
 * @param {number|string} crucesSTD
 * @return {string} "DENTRO DE RANGO" | "FUERA DE RANGO" | ""
 */
function calcularEstadoCruce(cruceReal, crucesSTD) {
  cruceReal = Number(cruceReal);
  crucesSTD = Number(crucesSTD);

  if (isNaN(cruceReal) || isNaN(crucesSTD) || crucesSTD <= 0) return "";

  const minimo = crucesSTD * (1 - MARGEN_TOLERANCIA_CRUCE);
  const maximo = crucesSTD * (1 + MARGEN_TOLERANCIA_CRUCE);

  return cruceReal >= minimo && cruceReal <= maximo
    ? "DENTRO DE RANGO"
    : "FUERA DE RANGO";
}

//----------------------------------
// VALIDAR REGISTROS
//----------------------------------
/**
 * Fila especial cuyo producto (CABO) solo puede fabricarse en sus
 * máquinas. Esa fila solo tiene 2 trenzadoras, por lo que se acepta
 * registrar 1 o 2 mediciones en vez de las 3 habituales.
 */
const FILA_DRIZA = "DRIZA";

/**
 * Valida, en el servidor, que se hayan enviado registros completos
 * (fila, número y cruce real) y que ninguna máquina (combinación
 * Fila + Número) esté repetida. Es una segunda capa de validación
 * además de la del formulario.
 *
 * Cantidad exigida:
 * - Si alguna medición es de la fila DRIZA: se aceptan 1 o 2
 *   registros (esa fila solo tiene 2 máquinas).
 * - En cualquier otro caso: se exigen exactamente
 *   CANTIDAD_MEDICIONES registros.
 *
 * Las filas SÍ pueden repetirse entre mediciones (dos o tres
 * mediciones pueden ser de la misma fila); lo que no puede
 * repetirse es la máquina exacta (Fila + Número).
 * @param {Array<{fila: string, numero: string, cruceReal: string|number}>} registros
 * @return {boolean}
 */
function registrosValidos(registros) {
  if (!Array.isArray(registros) || registros.length === 0) return false;
  if (registros.length > CANTIDAD_MEDICIONES) return false;

  const normalizados = registros.map((r) => ({
    fila: String(r.fila || "")
      .trim()
      .toUpperCase(),
    numero: String(r.numero || "")
      .trim()
      .toUpperCase(),
    cruceReal: r.cruceReal,
  }));

  // Fila, Número y Cruce Real deben estar completos en cada registro
  const completos = normalizados.every(
    (r) =>
      r.fila !== "" &&
      r.numero !== "" &&
      r.cruceReal !== "" &&
      r.cruceReal !== null &&
      r.cruceReal !== undefined &&
      !isNaN(Number(r.cruceReal)),
  );
  if (!completos) return false;

  const esModoDriza = normalizados.some((r) => r.fila === FILA_DRIZA);

  if (esModoDriza) {
    // La fila DRIZA solo tiene 2 máquinas: se acepta 1 o 2 mediciones
    if (normalizados.length < 1 || normalizados.length > 2) return false;
  } else {
    // Para el resto de filas se exigen exactamente CANTIDAD_MEDICIONES
    if (normalizados.length !== CANTIDAD_MEDICIONES) return false;
  }

  // No se puede repetir la misma máquina (Fila + Número), aunque
  // sí se pueden repetir filas entre sí.
  const claves = normalizados.map((r) => r.fila + "|" + r.numero);
  if (new Set(claves).size !== claves.length) return false;

  return true;
}

//----------------------------------
// GUARDAR REGISTRO
//----------------------------------
/**
 * Guarda un lote de mediciones (3 trenzadoras) en la hoja de
 * registros. Crea los encabezados si la hoja está vacía.
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
      "Supervisor",
      "Orden Trabajo",
      "Nombre Producto",
      "Cruces STD",
      "Fila",
      "Número",
      "Cruce Real",
      "Estado",
      "Observación",
    ]);
  }

  if (!registrosValidos(datos.registros)) {
    return {
      ok: false,
      mensaje:
        "Debe registrar " +
        CANTIDAD_MEDICIONES +
        " trenzadoras (o 1-2 si son de la fila " +
        FILA_DRIZA +
        "), con Fila, N° Trenzadora y Cruce Real completos, sin repetir la misma máquina (Fila + Número).",
    };
  }

  const info = obtenerFechaHora();

  const filasNuevas = datos.registros.map((registro) => [
    info.fecha,
    info.hora,
    info.turno,
    datos.codigo,
    datos.nombre,
    datos.ordenTrabajo,
    datos.nombreProducto,
    Number(datos.crucesSTD),
    registro.fila,
    registro.numero,
    Number(registro.cruceReal),
    calcularEstadoCruce(registro.cruceReal, datos.crucesSTD),
    String(registro.observacion || "")
      .trim()
      .toUpperCase(),
  ]);

  hoja
    .getRange(
      hoja.getLastRow() + 1,
      1,
      filasNuevas.length,
      filasNuevas[0].length,
    )
    .setValues(filasNuevas);

  return { ok: true, mensaje: "Registro guardado correctamente." };
}
