// ═══════════════════════════════════════════════════════════
//  CONFIGURACIÓN GLOBAL
// ═══════════════════════════════════════════════════════════
var CLOUD_NAME   = 'detm6isik';
var CLOUD_KEY    = '421275272827455';
var CLOUD_SECRET = '5xPue4Z3a2jbrKlc3Vr9pjQUbZw';
var REPLICATE    = 'r8_CdamyXgHwwvRzyKFXyZWVQUp97Vx3b031uH8y';
var SHEET_ID     = '1Duz5y1or2s3VimEZQH5JgBZfqzC3oGh9rlWwjJGR4VU';

function doGet(e) {
  if (e && e.parameter && e.parameter.pago === 'ok') {
    var cid = e.parameter.cid || '';
    var n   = parseInt(e.parameter.intentos || '0');
    if (cid && n > 0) agregarIntentos(cid, n);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Mi Figurita 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════
//  doPost — recibe llamadas fetch() desde mifiguritamundial.lat
//  Reemplaza google.script.run para funcionar fuera de GAS
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  var corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    var body   = JSON.parse(e.postData.contents);
    var fn     = body.fn;
    var result = {};

    if (fn === 'generarFigurita') {
      result = generarFigurita(body.pB64, body.fB64);
    } else if (fn === 'generarFiguritaRostroExacto') {
      result = generarFiguritaRostroExacto(body.pB64, body.fB64);
    } else if (fn === 'crearPagoMP') {
      result = crearPagoMP(body.plan, body.clientId);
    } else if (fn === 'obtenerOCrearCliente') {
      result = obtenerOCrearCliente(body.clientId);
    } else {
      result = { ok: false, error: 'Función no reconocida: ' + fn };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════
//  GOOGLE SHEETS — intentos por cliente
// ═══════════════════════════════════════════════════════
function obtenerOCrearCliente(clientId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == clientId) {
        return { ok: true, intentos: parseInt(data[i][1]) || 0 };
      }
    }
    var fila = data.length + 1;
    sheet.getRange(fila, 1).setValue(clientId);
    sheet.getRange(fila, 2).setValue(3);
    sheet.getRange(fila, 3).setValue(new Date().toISOString());
    return { ok: true, intentos: 3 };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function agregarIntentos(clientId, cantidad) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == clientId) {
        var actuales = parseInt(data[i][1]) || 0;
        sheet.getRange(i + 1, 2).setValue(actuales + cantidad);
        return { ok: true };
      }
    }
    var fila = data.length + 1;
    sheet.getRange(fila, 1).setValue(clientId);
    sheet.getRange(fila, 2).setValue(cantidad);
    sheet.getRange(fila, 3).setValue(new Date().toISOString());
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  MERCADOPAGO — Crear preferencia de pago (Checkout Pro)
//  Recibe el plan elegido y devuelve la URL de pago de MP
// ═══════════════════════════════════════════════════════════
var MP_ACCESS_TOKEN = 'APP_USR-5290599878424600-052320-9164823cbc30f6023603323a47bdb774-543316500';
var APP_URL = 'https://script.google.com/macros/s/AKfycbw-hTtUaXM32MmtBfMSRFH7wJrnnYVAsk35aNqlr7XLuXDw1pWulXwyZXbuxhpLsl4tmA/exec';

function crearPagoMP(plan, clientId) {
  try {
    var planes = {
      basico:   { titulo: 'Plan Básico - 1 Figurita',    precio: 1,  intentos: 1  },
      estrella: { titulo: 'Plan Estrella - 5 Figuritas', precio: 6,  intentos: 5  },
      album:    { titulo: 'Plan Álbum - 15 Figuritas',   precio: 15, intentos: 15 }
    };

    var p = planes[plan];
    if (!p) return { ok: false, error: 'Plan inválido' };

    var preference = {
      items: [{
        title:      p.titulo,
        quantity:   1,
        unit_price: p.precio,
        currency_id: 'USD'
      }],
      back_urls: {
        success: APP_URL + '?pago=ok&plan=' + plan + '&intentos=' + p.intentos + '&cid=' + clientId,
        failure: APP_URL + '?pago=error',
        pending: APP_URL + '?pago=pending'
      },
      auto_return: 'approved',
      statement_descriptor: 'FIGURITAS 2026'
    };

    var res = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + MP_ACCESS_TOKEN,
        'Content-Type':  'application/json'
      },
      payload:            JSON.stringify(preference),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var data = JSON.parse(res.getContentText());

    Logger.log('MP status: ' + code);

    if (code !== 201) {
      return { ok: false, error: 'MP error: ' + JSON.stringify(data) };
    }

    return { ok: true, url: data.init_point, id: data.id };

  } catch(e) {
    return { ok: false, error: 'MP excepción: ' + e.toString() };
  }
}



// ═══════════════════════════════════════════════════════════
//  BOTÓN 1 — IA COMPLETA
//  Face swap que transforma el rostro adoptando la estética
//  del jugador real de la figurita (mismo modelo que antes)
// ═══════════════════════════════════════════════════════════
function generarFigurita(fotoBase64, marcaBase64) {
  try {
    var fotoClean  = limpiarBase64(fotoBase64);
    var marcaClean = limpiarBase64(marcaBase64);

    // 1. Subir ambas imágenes a Cloudinary
    var urlFoto  = subirCloudinary(fotoClean);
    if (!urlFoto.ok)  return { ok: false, error: 'Cloudinary foto: '  + urlFoto.error };

    var urlMarca = subirCloudinary(marcaClean);
    if (!urlMarca.ok) return { ok: false, error: 'Cloudinary marco: ' + urlMarca.error };

    // 2. Face swap completo con IA (modelo original)
    //    input_image = marco (la figurita con el jugador)
    //    swap_image  = foto del cliente (su cara reemplaza la del jugador)
    var resultado = llamarReplicate({
      version: "d1d6ea8c8be89d664a07a457526f7128109dee7030fdac424788d762c71ed111",
      input: {
        input_image: urlMarca.url,
        swap_image:  urlFoto.url
      }
    });

    return resultado;

  } catch(e) {
    return { ok: false, error: 'Excepción Botón 1: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  REMOVE-BG — usado por el Botón 2
//  Usa Remove.bg API — calidad profesional de recorte
//  Devuelve base64 PNG con fondo transparente
// ═══════════════════════════════════════════════════════════
function removeBgFoto(fotoBase64) {
  var REMOVEBG_KEY = 'UTBStFgBCewBpGmQLYkctnYJ';

  try {
    var fotoClean = limpiarBase64(fotoBase64);

    // Subir foto a Cloudinary primero para obtener URL pública
    var uploadResult = subirCloudinary(fotoClean);
    if (!uploadResult.ok) return { ok: false, error: 'Cloudinary: ' + uploadResult.error };

    // Llamar Remove.bg via URL — más confiable que base64
    var res = UrlFetchApp.fetch('https://api.remove.bg/v1.0/removebg', {
      method: 'post',
      headers: {
        'X-Api-Key': REMOVEBG_KEY
      },
      payload: {
        'image_url': uploadResult.url,  // URL pública en lugar de base64
        'size':      'auto',
        'format':    'png',
        'type':      'person'
      },
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();

    // Log para debug
    Logger.log('Remove.bg status: ' + code);

    if (code !== 200) {
      var err = res.getContentText();
      Logger.log('Remove.bg error body: ' + err);
      return { ok: false, error: 'Remove.bg error ' + code + ': ' + err };
    }

    // Verificar que la respuesta tiene contenido
    var bytes = res.getContent();
    if (!bytes || bytes.length < 100) {
      return { ok: false, error: 'Remove.bg devolvió respuesta vacía' };
    }

    Logger.log('Remove.bg OK - bytes: ' + bytes.length);

    var b64     = Utilities.base64Encode(bytes);
    var dataUrl = 'data:image/png;base64,' + b64;

    return { ok: true, url: dataUrl };

  } catch(e) {
    return { ok: false, error: 'removeBg: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  UPSCALE — mejora calidad de imagen 4x sin alterar nada
//  Usa real-esrgan en Replicate — solo agrega píxeles
//  La imagen queda idéntica pero mucho más nítida
// ═══════════════════════════════════════════════════════════
function upscaleFigurita(imagenBase64) {
  try {
    var limpio = limpiarBase64(imagenBase64);

    // Subir a Cloudinary
    var upload = subirCloudinary(limpio);
    if (!upload.ok) return { ok: false, error: 'Cloudinary upscale: ' + upload.error };

    // Llamar real-esrgan — upscale 4x sin modificar la imagen
    var resultado = llamarReplicate({
      version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
      input: {
        image:  upload.url,
        scale:  2,           // 2x es suficiente y más rápido que 4x
        face_enhance: false  // MUY IMPORTANTE — false para no alterar la cara
      }
    });

    return resultado;

  } catch(e) {
    return { ok: false, error: 'upscale: ' + e.toString() };
  }
}




// ═══════════════════════════════════════════════════════════
//  BOTÓN 2 — ROSTRO EXACTO via Segmind faceswap-v5
//
//  Flujo:
//    1. Subir foto cliente y marco a Cloudinary → URLs públicas
//    2. Llamar Segmind faceswap-v5
//       source_image = foto del cliente
//       target_image = figurita
//    → Intercambio ultrarrápido de rostros preservando identidad
// ═══════════════════════════════════════════════════════════
function generarFiguritaRostroExacto(fotoBase64, marcaBase64) {
  var SEGMIND_KEY = 'SG_62105cca1900f386';

  try {
    var fotoClean  = limpiarBase64(fotoBase64);
    var marcaClean = limpiarBase64(marcaBase64);

    // PASO 1 — Subir ambas imágenes a Cloudinary
    var urlFoto  = subirCloudinary(fotoClean);
    if (!urlFoto.ok)  return { ok: false, error: 'Cloudinary foto: ' + urlFoto.error };

    var urlMarca = subirCloudinary(marcaClean);
    if (!urlMarca.ok) return { ok: false, error: 'Cloudinary marco: ' + urlMarca.error };

    // PASO 2 — Segmind faceswap-v5
    var payload = {
      source_image: urlFoto.url,   // cara del cliente
      target_image: urlMarca.url,  // figurita destino
      image_format: 'png',
      quality:      95
    };

    var res = UrlFetchApp.fetch('https://api.segmind.com/v1/faceswap-v5', {
      method: 'post',
      headers: {
        'x-api-key':    SEGMIND_KEY,
        'Content-Type': 'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    Logger.log('Segmind status: ' + code);

    if (code !== 200) {
      var err = res.getContentText();
      Logger.log('Segmind error: ' + err);
      return { ok: false, error: 'Segmind error ' + code + ': ' + err };
    }

    // Segmind devuelve la imagen directamente como bytes
    var bytes  = res.getContent();
    var b64    = Utilities.base64Encode(bytes);
    var imgUrl = 'data:image/png;base64,' + b64;

    // Subir resultado a Cloudinary para obtener URL pública
    var upload = subirCloudinary(b64);
    if (!upload.ok) return { ok: false, error: 'Cloudinary resultado: ' + upload.error };

    return { ok: true, url: upload.url };

  } catch(e) {
    return { ok: false, error: 'Excepción Segmind: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  REMOVE-BG — usado por el Botón 2
//  Usa Remove.bg API — calidad profesional de recorte
//  Devuelve base64 PNG con fondo transparente
// ═══════════════════════════════════════════════════════════
function removeBgFoto(fotoBase64) {
  var REMOVEBG_KEY = 'UTBStFgBCewBpGmQLYkctnYJ';

  try {
    var fotoClean = limpiarBase64(fotoBase64);

    // Subir foto a Cloudinary primero para obtener URL pública
    var uploadResult = subirCloudinary(fotoClean);
    if (!uploadResult.ok) return { ok: false, error: 'Cloudinary: ' + uploadResult.error };

    // Llamar Remove.bg via URL — más confiable que base64
    var res = UrlFetchApp.fetch('https://api.remove.bg/v1.0/removebg', {
      method: 'post',
      headers: {
        'X-Api-Key': REMOVEBG_KEY
      },
      payload: {
        'image_url': uploadResult.url,  // URL pública en lugar de base64
        'size':      'auto',
        'format':    'png',
        'type':      'person'
      },
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();

    // Log para debug
    Logger.log('Remove.bg status: ' + code);

    if (code !== 200) {
      var err = res.getContentText();
      Logger.log('Remove.bg error body: ' + err);
      return { ok: false, error: 'Remove.bg error ' + code + ': ' + err };
    }

    // Verificar que la respuesta tiene contenido
    var bytes = res.getContent();
    if (!bytes || bytes.length < 100) {
      return { ok: false, error: 'Remove.bg devolvió respuesta vacía' };
    }

    Logger.log('Remove.bg OK - bytes: ' + bytes.length);

    var b64     = Utilities.base64Encode(bytes);
    var dataUrl = 'data:image/png;base64,' + b64;

    return { ok: true, url: dataUrl };

  } catch(e) {
    return { ok: false, error: 'removeBg: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  UPSCALE — mejora calidad de imagen 4x sin alterar nada
//  Usa real-esrgan en Replicate — solo agrega píxeles
//  La imagen queda idéntica pero mucho más nítida
// ═══════════════════════════════════════════════════════════
function upscaleFigurita(imagenBase64) {
  try {
    var limpio = limpiarBase64(imagenBase64);

    // Subir a Cloudinary
    var upload = subirCloudinary(limpio);
    if (!upload.ok) return { ok: false, error: 'Cloudinary upscale: ' + upload.error };

    // Llamar real-esrgan — upscale 4x sin modificar la imagen
    var resultado = llamarReplicate({
      version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
      input: {
        image:  upload.url,
        scale:  2,           // 2x es suficiente y más rápido que 4x
        face_enhance: false  // MUY IMPORTANTE — false para no alterar la cara
      }
    });

    return resultado;

  } catch(e) {
    return { ok: false, error: 'upscale: ' + e.toString() };
  }
}





// ═══════════════════════════════════════════════════════════
//  REMOVE-BG — usado por el Botón 2
//  Usa Remove.bg API — calidad profesional de recorte
//  Devuelve base64 PNG con fondo transparente
// ═══════════════════════════════════════════════════════════
function removeBgFoto(fotoBase64) {
  var REMOVEBG_KEY = 'UTBStFgBCewBpGmQLYkctnYJ';

  try {
    var fotoClean = limpiarBase64(fotoBase64);

    // Subir foto a Cloudinary primero para obtener URL pública
    var uploadResult = subirCloudinary(fotoClean);
    if (!uploadResult.ok) return { ok: false, error: 'Cloudinary: ' + uploadResult.error };

    // Llamar Remove.bg via URL — más confiable que base64
    var res = UrlFetchApp.fetch('https://api.remove.bg/v1.0/removebg', {
      method: 'post',
      headers: {
        'X-Api-Key': REMOVEBG_KEY
      },
      payload: {
        'image_url': uploadResult.url,  // URL pública en lugar de base64
        'size':      'auto',
        'format':    'png',
        'type':      'person'
      },
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();

    // Log para debug
    Logger.log('Remove.bg status: ' + code);

    if (code !== 200) {
      var err = res.getContentText();
      Logger.log('Remove.bg error body: ' + err);
      return { ok: false, error: 'Remove.bg error ' + code + ': ' + err };
    }

    // Verificar que la respuesta tiene contenido
    var bytes = res.getContent();
    if (!bytes || bytes.length < 100) {
      return { ok: false, error: 'Remove.bg devolvió respuesta vacía' };
    }

    Logger.log('Remove.bg OK - bytes: ' + bytes.length);

    var b64     = Utilities.base64Encode(bytes);
    var dataUrl = 'data:image/png;base64,' + b64;

    return { ok: true, url: dataUrl };

  } catch(e) {
    return { ok: false, error: 'removeBg: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════
//  UPSCALE — mejora calidad de imagen 4x sin alterar nada
//  Usa real-esrgan en Replicate — solo agrega píxeles
//  La imagen queda idéntica pero mucho más nítida
// ═══════════════════════════════════════════════════════════
function upscaleFigurita(imagenBase64) {
  try {
    var limpio = limpiarBase64(imagenBase64);

    // Subir a Cloudinary
    var upload = subirCloudinary(limpio);
    if (!upload.ok) return { ok: false, error: 'Cloudinary upscale: ' + upload.error };

    // Llamar real-esrgan — upscale 4x sin modificar la imagen
    var resultado = llamarReplicate({
      version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
      input: {
        image:  upload.url,
        scale:  2,           // 2x es suficiente y más rápido que 4x
        face_enhance: false  // MUY IMPORTANTE — false para no alterar la cara
      }
    });

    return resultado;

  } catch(e) {
    return { ok: false, error: 'upscale: ' + e.toString() };
  }
}




// ═══════════════════════════════════════════════════════════
//  BOTÓN 2 — ROSTRO EXACTO via Segmind faceswap-v5
//
//  Flujo:
//    1. Subir foto cliente y marco a Cloudinary → URLs públicas
//    2. Llamar Segmind faceswap-v5
//       source_image = foto del cliente
//       target_image = figurita
//    → Intercambio ultrarrápido de rostros preservando identidad



// ═══════════════════════════════════════════════════════════
//  LLAMAR REPLICATE — con polling integrado
// ═══════════════════════════════════════════════════════════
function llamarReplicate(payload) {
  var MAX_REINTENTOS = 3;      // hasta 3 intentos automáticos
  var ESPERA_RATE    = 20000;  // 20 segundos si hay rate limit (429)

  for (var intento = 1; intento <= MAX_REINTENTOS; intento++) {

    var res = UrlFetchApp.fetch('https://api.replicate.com/v1/predictions', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + REPLICATE,
        'Content-Type':  'application/json',
        'Prefer':        'wait=60'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();

    // Rate limit (429) — esperar y reintentar automáticamente
    if (code === 429) {
      if (intento < MAX_REINTENTOS) {
        Utilities.sleep(ESPERA_RATE);
        continue; // volver a intentar
      } else {
        return { ok: false, error: 'Replicate ocupado, intentá en unos segundos.' };
      }
    }

    var pred = JSON.parse(res.getContentText());

    // Respuesta inmediata (Prefer: wait=60 la resuelve en el momento)
    if (pred.status === 'succeeded' && pred.output) {
      var out = Array.isArray(pred.output) ? pred.output[0] : pred.output;
      return { ok: true, url: out };
    }

    // Si no resolvió de inmediato, hacer polling
    var id = pred.id;
    if (!id) return { ok: false, error: 'Replicate no devolvió ID: ' + JSON.stringify(pred) };

    for (var i = 0; i < 30; i++) {
      Utilities.sleep(2000);
      var poll = JSON.parse(UrlFetchApp.fetch(
        'https://api.replicate.com/v1/predictions/' + id,
        {
          headers:            { 'Authorization': 'Bearer ' + REPLICATE },
          muteHttpExceptions: true
        }
      ).getContentText());

      if (poll.status === 'succeeded' && poll.output) {
        var url = Array.isArray(poll.output) ? poll.output[0] : poll.output;
        return { ok: true, url: url };
      }
      if (poll.status === 'failed') {
        return { ok: false, error: 'Replicate falló: ' + (poll.error || 'error desconocido') };
      }
    }

    return { ok: false, error: 'Timeout — el proceso tardó demasiado, intentá de nuevo' };
  }

  return { ok: false, error: 'No se pudo conectar con Replicate después de varios intentos.' };
}


// ═══════════════════════════════════════════════════════════
//  SUBIR A CLOUDINARY
// ═══════════════════════════════════════════════════════════
function subirCloudinary(base64clean) {
  try {
    var timestamp = Math.floor(Date.now() / 1000).toString();
    var toSign    = 'timestamp=' + timestamp + CLOUD_SECRET;
    var signature = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_1,
      toSign,
      Utilities.Charset.UTF_8
    ).map(function(b) {
      return ('0' + (b & 0xFF).toString(16)).slice(-2);
    }).join('');

    var res = UrlFetchApp.fetch(
      'https://api.cloudinary.com/v1_1/' + CLOUD_NAME + '/image/upload',
      {
        method:  'post',
        payload: {
          file:      'data:image/jpeg;base64,' + base64clean,
          api_key:   CLOUD_KEY,
          timestamp: timestamp,
          signature: signature
        },
        muteHttpExceptions: true
      }
    );

    var data = JSON.parse(res.getContentText());
    if (data.secure_url) return { ok: true, url: data.secure_url };
    return { ok: false, error: JSON.stringify(data) };

  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}


// ═══════════════════════════════════════════════════════════
//  UTILIDADES
// ═══════════════════════════════════════════════════════════
function limpiarBase64(b64) {
  return b64.indexOf('data:') === 0 ? b64.split(',')[1] : b64;
}
