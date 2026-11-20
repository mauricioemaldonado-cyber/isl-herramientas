const SPREADSHEET_ID = "1y0dpFWgiPL1zIjF2xP4WP-u7w4D66rMrIWyCrfRkDnw";

// ── VERIFICAR ACCESO ─────────────────────────────────────────────────────────
// Lee la hoja "Accesos" y devuelve si el email tiene permiso para la app dada.
// La hoja Accesos debe tener en fila 1 los encabezados:
//   Email | Nombre | Herramientas | Manuales | Remitos | EPP | Extintores | Izaje | Taller | Compras | Pedidos | Mtto | Hub | HyS | Seguim | Stock
function verificarAcceso(email, appKey) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Accesos");
    if (!sheet) return { ok: true, autorizado: false, error: "Hoja Accesos no encontrada" };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, autorizado: false };

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var emailCol = headers.indexOf("Email");
    var appCol = headers.indexOf(appKey);

    if (emailCol === -1) return { ok: true, autorizado: false, error: "Columna Email no encontrada" };
    if (appCol === -1) return { ok: true, autorizado: false, error: "App " + appKey + " no encontrada en Accesos" };

    var emailLower = String(email).toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][emailCol] || "").toLowerCase().trim();
      if (rowEmail === emailLower) {
        var val = String(data[i][appCol] || "").trim();
        var autorizado = (val === "✓" || val === "1" || val.toLowerCase() === "si" || val.toLowerCase() === "sí" || val === "TRUE" || val === "true");
        return { ok: true, autorizado: autorizado };
      }
    }
    return { ok: true, autorizado: false };
  } catch(err) {
    return { ok: false, autorizado: false, error: err.message };
  }
}

// ── DO GET ───────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Verificar acceso
    if (e.parameter.action === "verificarAcceso") {
      var email  = e.parameter.email  || "";
      var appKey = e.parameter.app    || "";
      var result = verificarAcceso(email, appKey);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.payload) {
      var body = JSON.parse(e.parameter.payload);
      if (body.action === "append") {
        var sheet = ss.getSheetByName(body.sheet);
        if (!sheet) throw new Error("Hoja no encontrada: " + body.sheet);
        sheet.getRange(sheet.getLastRow()+1, 1, body.values.length, body.values[0].length)
             .setValues(body.values);
        return ContentService.createTextOutput(JSON.stringify({ok:true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (body.action === "update") {
        ss.getRange(body.range).setValues(body.values);
        return ContentService.createTextOutput(JSON.stringify({ok:true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    var range = e.parameter.range;
    if (!range) throw new Error("Falta parametro: range");
    var values = ss.getRange(range).getValues();
    var result = JSON.stringify({ok:true, values:values});

    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + "(" + result + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService.createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── DO POST ──────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (body.action === "append") {
      var sheet = ss.getSheetByName(body.sheet);
      if (!sheet) throw new Error("Hoja no encontrada: " + body.sheet);
      sheet.getRange(sheet.getLastRow()+1, 1, body.values.length, body.values[0].length)
           .setValues(body.values);
    } else if (body.action === "update") {
      ss.getRange(body.range).setValues(body.values);
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── SETUP HOJA ACCESOS ───────────────────────────────────────────────────────
// Ejecutar UNA VEZ manualmente para crear la hoja Accesos con los usuarios.
function setupHojaAccesos() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Accesos");
  if (!sheet) {
    sheet = ss.insertSheet("Accesos");
  } else {
    sheet.clearContents();
  }

  var headers = ["Email","Nombre","Herramientas","Manuales","Remitos","EPP","Extintores","Izaje","Taller","Compras","Pedidos","Mtto","Hub","HyS","Seguim","Stock"];

  // ✓ = autorizado, vacío = sin acceso
  var v = "✓";
  var datos = [
    headers,
    // Email                              Nombre              Her  Man  Rem  EPP  Ext  Izaje Tall Comp Ped  Mtto Hub  HyS  Seg  Stock
    ["mauricio.e.maldonado@gmail.com", "MAURICIO MALDONADO",  v,   v,   v,   v,   v,   v,   v,   v,   v,   v,   v,   v,   v,   v  ],
    ["islproyecto2024@gmail.com",      "DAVID TESIDO",        v,   v,   v,   v,   v,   v,   v,   "",  v,   v,   v,   v,   "",  v  ],
    ["rledesma164@gmail.com",          "RAMON LEDESMA",       "",  "",  v,   "",  "",  "",  "",  "",  v,   "",  "",  "",  "",  ""  ],
    ["yayodanielgomez@gmail.com",      "DANIEL GOMEZ",        "",  "",  v,   "",  "",  "",  v,   "",  v,   "",  "",  "",  "",  v   ],
    ["instrumentalsanlzo01@gmail.com", "ANGEL GOMEZ",         "",  "",  v,   "",  "",  "",  v,   "",  v,   "",  "",  "",  "",  ""  ],
    ["margiecampos37@gmail.com",       "LUIS HERNANDEZ",      "",  "",  v,   "",  "",  "",  "",  "",  v,   "",  "",  "",  v,   v   ],
    ["j.loaisa13@gmail.com",           "JOSE LOAISA",         "",  "",  v,   "",  "",  "",  "",  v,   v,   "",  "",  "",  v,   v   ],
    ["sebastianalthaus5@gmail.com",    "SEBASTIAN ALTHAUS",   "",  "",  v,   "",  "",  "",  "",  "",  v,   "",  "",  "",  v,   v   ],
    // César Oliva — falta email, solo Mtto por ahora
    // ["cesar@gmail.com",             "CESAR OLIVA",         "",  "",  "",  "",  "",  "",  "",  "",  "",  v,   "",  "",  "",  ""  ],
  ];

  sheet.getRange(1, 1, datos.length, headers.length).setValues(datos);

  // Formato encabezado
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1A3A5C").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("✓ Hoja Accesos creada con " + (datos.length-1) + " usuarios");
}

// ── HELPERS (sin cambios) ────────────────────────────────────────────────────
function listarHojas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojas = ss.getSheets().map(s => s.getName());
  Logger.log(hojas.join(", "));
  return hojas;
}

function setupEncabezados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const hm = ss.getSheetByName("HerramientasManuales");
  hm.getRange(1,1,1,17).setValues([["ID","Codigo","Descripcion","Categoria","Marca","Modelo","NroSerie","TipoTenencia","ValorCompra","FechaAlta","NodoActual","Estado","Proveedor","NroFactura","FechaFactura","Observaciones","EsExistente"]]);
  hm.getRange(1,1,1,17).setBackground("#1A3A5C").setFontColor("#FFFFFF").setFontWeight("bold");
  hm.setFrozenRows(1);

  const mm = ss.getSheetByName("MovimientosManuales");
  mm.getRange(1,1,1,12).setValues([["ID","Fecha","Hora","TipoMovimiento","CodigoHerramienta","Descripcion","NodoOrigen","NodoDestino","Responsable","Receptor","NroDocumento","Observaciones"]]);
  mm.getRange(1,1,1,12).setBackground("#1A3A5C").setFontColor("#FFFFFF").setFontWeight("bold");
  mm.setFrozenRows(1);

  const am = ss.getSheetByName("ActasManuales");
  am.getRange(1,1,1,16).setValues([["ID","NroActa","Tipo","Fecha","CodigoHerramienta","Descripcion","ReceptorNombre","ReceptorDNI","ReceptorLegajo","ReceptorRol","Obra","ValorReposicion","EstadoEntrega","FechaDevolucion","EstadoDevolucion","Observaciones"]]);
  am.getRange(1,1,1,16).setBackground("#1A3A5C").setFontColor("#FFFFFF").setFontWeight("bold");
  am.setFrozenRows(1);

  Logger.log("✓ Encabezados creados en las 3 hojas");
}
