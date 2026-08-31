/**
 * Función principal que sirve la aplicación web
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ERP DaLiMar')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función auxiliar para incluir archivos HTML dentro de otros (como componentes)
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename)
      .evaluate()
      .getContent();
}

// ==========================================
//   CONTROL DE CONCURRENCIA (LockService)
// ==========================================
//
// Todas las funciones que ESCRIBEN en la hoja (crear, editar, eliminar,
// registrar) pasan su lógica por este helper. Evita que dos usuarios
// trabajando al mismo tiempo (en distintas computadoras) corrompan datos
// por escribir sobre la misma fila o duplicar registros sin querer.
//
// Si no se puede obtener el lock en 10 segundos (otro usuario esta
// guardando algo justo en ese instante), se devuelve un error amigable
// en vez de dejar al usuario esperando indefinidamente o fallar en silencio.
function ejecutarConLock(funcionLogica) {
  const lock = LockService.getScriptLock();
  try {
    const obtenido = lock.tryLock(10000); // espera hasta 10 segundos
    if (!obtenido) {
      return { exito: false, mensaje: "⚠️ El sistema está ocupado guardando otro cambio. Intenta de nuevo en unos segundos." };
    }
    return funcionLogica();
  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor: " + error.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
//        MÓDULO: DIRECTORIO DE CLIENTES
// ==========================================

function registrarCliente(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Clientes");

    if (!hoja) {
      hoja = ss.insertSheet("Clientes");
      hoja.appendRow(["Tipo", "Nombre", "DUI/NIT", "NRC", "Correo", "Telefono", "Celular", "Direccion"]);
    }

    const db = hoja.getDataRange().getValues();

    if (datos.esEdicion) {
      let filaDestino = -1;
      for (let i = 1; i < db.length; i++) {
        if (db[i][2].toString() === datos.duiNitOriginal.toString()) {
          filaDestino = i + 1;
          break;
        }
      }

      if (filaDestino !== -1) {
        hoja.getRange(filaDestino, 1, 1, 8).setValues([[
          datos.tipo,
          datos.nombre,
          "'" + datos.duiNit,
          "'" + datos.nrc,
          datos.correo,
          datos.telefono,
          datos.celular,
          datos.direccion
        ]]);
        return { exito: true, mensaje: "¡Cliente actualizado correctamente con éxito!" };
      } else {
        return { exito: false, mensaje: "Error: No se encontró el registro original para modificar." };
      }

    } else {
      for (let i = 1; i < db.length; i++) {
        if (db[i][2].toString() === datos.duiNit.toString()) {
          return { exito: false, mensaje: "⚠️ Ya existe un cliente registrado con ese DUI o NIT." };
        }
      }

      hoja.appendRow([
        datos.tipo,
        datos.nombre,
        "'" + datos.duiNit,
        "'" + datos.nrc,
        datos.correo,
        datos.telefono,
        datos.celular,
        datos.direccion
      ]);

      return { exito: true, mensaje: "¡Cliente guardado en la base de datos de manera exitosa!" };
    }

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor: " + error.message };
  }
}

function obtenerClientesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Clientes");
    if (!hoja) return [];

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    let clientes = [];
    for (let i = 1; i < datos.length; i++) {
      clientes.push({
        tipo: datos[i][0],
        nombre: datos[i][1],
        duiNit: datos[i][2],
        nrc: datos[i][3],
        correo: datos[i][4],
        telefono: datos[i][5],
        celular: datos[i][6],
        direccion: datos[i][7]
      });
    }
    return clientes;
  } catch (e) {
    return [];
  }
}

function eliminarClienteServidor(duiNit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Clientes");
    if (!hoja) return { exito: false, mensaje: "No se encontró la hoja de Clientes." };

    const db = hoja.getDataRange().getValues();
    let filaAEliminar = -1;

    for (let i = 1; i < db.length; i++) {
      if (db[i][2].toString() === duiNit.toString()) {
        filaAEliminar = i + 1;
        break;
      }
    }

    if (filaAEliminar !== -1) {
      hoja.deleteRow(filaAEliminar);
      return { exito: true, mensaje: "¡Cliente eliminado de la base de datos correctamente!" };
    } else {
      return { exito: false, mensaje: "No se encontró el registro para eliminar." };
    }
  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor: " + error.message };
  }
}

// ==========================================
//          MÓDULO: GESTIÓN DE VENTAS
// ==========================================
// NOTA: ModVentas.html envía un payload con esta forma:
// { tipoOperacion, cliente: {...}, artículos: [...], financiero: {...}, formaPago, fechaRegistro }

/**
 * Procesa un Pedido Directo -> va directo a producción
 */
/**
 * Procesa un Pedido Directo -> va directo a producción
 * NOTA: Un pedido a producción consume INSUMOS/materia prima para fabricar,
 * no descuenta directamente el producto terminado. Esa lógica de consumo de
 * insumos pertenece al futuro módulo de Producción, no se implementa aquí.
 */
function procesarPedidoProduccionServidor(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaProd = ss.getSheetByName("Pedidos_Produccion");
    if (!hojaProd) {
      hojaProd = ss.insertSheet("Pedidos_Produccion");
      hojaProd.appendRow(["ID_Pedido", "Fecha", "Cliente", "DUI_NIT", "Tipo_Doc", "Items_Detalle", "Subtotal", "IVA", "Total", "Estado", "Notas"]);
    }

    const cliente = payload.cliente || {};
    const anioP = new Date().getFullYear();
    const dbP = hojaProd.getDataRange().getValues();
    let maxNumP = 0;
    for (let i = 1; i < dbP.length; i++) {
      const m = (dbP[i][0] || "").toString().match(/PED-\d{4}-(\d+)/);
      if (m) maxNumP = Math.max(maxNumP, parseInt(m[1]));
    }
    const idPedido = "PED-" + anioP + "-" + String(maxNumP + 1).padStart(4, "0");

    hojaProd.appendRow([
      idPedido,
      new Date(),
      cliente.nombre,
      cliente.duiNit,
      cliente.tipoCliente,
      JSON.stringify(payload.artículos || []),
      payload.financiero ? payload.financiero.subtotal : 0,
      payload.financiero ? payload.financiero.iva : 0,
      payload.financiero ? payload.financiero.total : 0,
      "Pendiente de Producción",
      payload.notas || ""
    ]);

    // Generar placeholder de orden/recibo para el cliente
    generarDocumentoFiscalServidor({
      id: idPedido,
      tipo: 'Orden_Produccion',
      tipoOperacion: 'Pedido a Producción',
      cliente: cliente,
      articulos: payload.artículos || [],
      financiero: payload.financiero,
      formaPago: 'Pendiente',
      notas: payload.notas || ''
    });

    // Cuenta por cobrar: el pedido se cobra cuando esté listo
    registrarIngresoServidor({
      referencia: idPedido,
      tipo: 'Cuenta por Cobrar',
      cliente: cliente.nombre || '',
      monto: payload.financiero ? payload.financiero.total : 0,
      formaPago: 'Pendiente de cobro'
    });

    return { exito: true, mensaje: '¡Pedido ' + idPedido + ' enviado a Producción!', id: idPedido };
  } catch (error) {
    return { exito: false, mensaje: "Error en pedido a producción: " + error.message };
  }
}

/**
 * Procesa una Cotización -> queda pendiente y se envía por correo si hay correo
 */
/**
 * Procesa una Cotización -> queda pendiente y se envía por correo si hay correo
 * NOTA: Una cotización es solo una propuesta, no una venta confirmada.
 * No descuenta inventario hasta que el cliente la apruebe y se convierta en pedido/venta.
 */
function procesarCotizacionEnvioServidor(payload) {
  return ejecutarConLock(function() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let hojaCot = ss.getSheetByName("Cotizaciones");
      if (!hojaCot) {
        hojaCot = ss.insertSheet("Cotizaciones");
        hojaCot.appendRow(["ID_Cotizacion", "Fecha", "Cliente", "DUI_NIT", "Tipo_Doc", "Correo", "Items_Detalle", "Subtotal", "IVA", "Total", "Estado", "Notas"]);
      }

      const db = hojaCot.getDataRange().getValues();
      const cliente = payload.cliente || {};
      const total = payload.financiero ? payload.financiero.total : 0;

      // Si viene un ID de cotización a editar, actualiza esa fila en vez de crear nueva
      const idEditando = (payload.idCotizacionEditando || "").toString().trim();
      if (idEditando) {
        for (let i = 1; i < db.length; i++) {
          if (db[i][0].toString() === idEditando) {
            hojaCot.getRange(i + 1, 1, 1, 12).setValues([[
              idEditando, db[i][1], // conserva la fecha original
              cliente.nombre || "", cliente.duiNit || "", cliente.tipoCliente || "", cliente.correo || "",
              JSON.stringify(payload.artículos || []),
              payload.financiero ? payload.financiero.subtotal : 0,
              payload.financiero ? payload.financiero.iva : 0,
              total, db[i][10] || "Pendiente de Aprobación", payload.notas || ""
            ]]);
            registrarAuditoria("Ventas", "Editar", "Cotizacion actualizada: " + idEditando);
            return { exito: true, mensaje: "¡Cotización " + idEditando + " actualizada!", id: idEditando };
          }
        }
      }

      // Correlativo nuevo
      const anio = new Date().getFullYear();
      let maxNum = 0;
      for (let i = 1; i < db.length; i++) {
        const idFila = (db[i][0] || "").toString();
        const match = idFila.match(/COT-\d{4}-(\d+)/);
        if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
      }
      const idCotizacion = "COT-" + anio + "-" + String(maxNum + 1).padStart(4, "0");

      hojaCot.appendRow([
        idCotizacion,
        new Date(),
        cliente.nombre || "",
        cliente.duiNit || "",
        cliente.tipoCliente || "",
        cliente.correo || "",
        JSON.stringify(payload.artículos || []),
        payload.financiero ? payload.financiero.subtotal : 0,
        payload.financiero ? payload.financiero.iva : 0,
        total,
        "Pendiente de Aprobación",
        payload.notas || ""
      ]);

      // El correo SOLO se envia si el frontend lo indica explicitamente
      // (payload.enviarCorreo === true), lo cual solo ocurre en modo PRO.
      // En modo Estandar, la cotizacion simplemente se registra y listo.
      if (payload.enviarCorreo && cliente.correo && cliente.correo.trim() !== "") {
        try {
          const empresa = obtenerConfiguracionEmpresaServidor() || {};
          const nombreEmpresa = empresa.nombreEmpresa || "Nuestra empresa";
          const asunto = "Cotización " + idCotizacion + " — " + nombreEmpresa;
          const lineasItems = (payload.artículos || []).map(it =>
            "• " + it.desc + " — Cant: " + it.cant + " @ $" + parseFloat(it.prec || 0).toFixed(2) + " = $" + parseFloat(it.total || 0).toFixed(2)
          ).join("\n");
          const cuerpo = "Estimado/a " + (cliente.nombre || "") + ",\n\n" +
            "Le compartimos el detalle de su cotización " + idCotizacion + ":\n\n" +
            lineasItems + "\n\n" +
            "Total: $" + total.toFixed(2) + "\n\n" +
            (payload.notas ? "Notas: " + payload.notas + "\n\n" : "") +
            "Quedamos a su disposición para cualquier consulta.\n\nAtentamente,\n" + nombreEmpresa;
          MailApp.sendEmail(cliente.correo, asunto, cuerpo);
        } catch (errMail) {
          console.warn("No se pudo enviar el correo: " + errMail.message);
        }
      }

      registrarAuditoria("Ventas", "Cotizar", "Cotizacion guardada: " + idCotizacion);
      const msg = payload.enviarCorreo
        ? "¡Cotización " + idCotizacion + " guardada y enviada por correo!"
        : "¡Cotización " + idCotizacion + " guardada correctamente!";
      return { exito: true, mensaje: msg, id: idCotizacion };
    } catch (error) {
      return { exito: false, mensaje: "Error al guardar la cotización: " + error.message };
    }
  });
}

// ==========================================
//      MÓDULO: VENTA DIRECTA (MOSTRADOR)
// ==========================================

/**
 * Procesa una Venta Directa: cobro inmediato, no entra a producción.
 * Requiere forma de pago. Descuenta stock solo si el producto tiene control de stock activo.
 * @param {Object} payload { cliente, artículos, financiero, formaPago, fechaRegistro }
 */
function procesarVentaDirectaServidor(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cliente = payload.cliente || {};
    const formaPago = payload.formaPago || "";
    const articulos = payload.artículos || [];

    if (!formaPago) {
      return { exito: false, mensaje: "Debe indicar la forma de pago para registrar la venta." };
    }

    // Separar ítems por tipo de producto: Producción vs Venta Directa/Servicio
    // El tipo lo determina el catálogo, no el usuario — así simplificamos el flujo.
    const mapaProductos = {};
    try {
      const prods = obtenerProductosServidor();
      prods.forEach(p => {
        (p.variantes || []).forEach(v => {
          mapaProductos[v.sku] = p.tipo || 'Producto';
        });
        // También indexar por nombre normalizado como fallback
        mapaProductos['desc:' + (p.nombre || '').toLowerCase().trim()] = p.tipo || 'Producto';
      });
    } catch(e) { console.warn("No se pudo obtener catálogo para separar tipos: " + e.message); }

    function tipoItem(item) {
      const sku = (item.sku || '').trim();
      if (sku && mapaProductos[sku]) return mapaProductos[sku];
      const desc = (item.desc || item.descripcion || '').toLowerCase().trim();
      return mapaProductos['desc:' + desc] || 'Producto';
    }

    const itemsProduccion = articulos.filter(it => tipoItem(it) === 'Producción');
    const itemsVentaDirecta = articulos.filter(it => tipoItem(it) !== 'Producción');

    // Validar stock solo para los ítems de venta directa
    if (itemsVentaDirecta.length > 0) {
      const validacion = validarStockDisponibleParaVenta(itemsVentaDirecta);
      if (!validacion.ok) return { exito: false, mensaje: validacion.mensaje };
      if (validacion.aviso && !payload.confirmarDespachoMultiBodega) {
        return { exito: false, aviso: true, mensajeAviso: validacion.mensajeAviso };
      }
    }

    const anio = new Date().getFullYear();
    const idsGenerados = [];
    let idVentaPrincipal = null;

    // Generar ID de orden padre que agrupa toda la operación
    // (aunque sea solo VD o solo PED, siempre existe el ORD- de referencia)
    let hojaOrdenes = ss.getSheetByName("Ordenes_Venta");
    if (!hojaOrdenes) {
      hojaOrdenes = ss.insertSheet("Ordenes_Venta");
      hojaOrdenes.appendRow(["ID_Orden","Fecha","Cliente","DUI_NIT","Items_JSON",
        "Subtotal","IVA","Total","Forma_Pago","Estado","Notas","IDs_Generados"]);
    }
    const dbOrd = hojaOrdenes.getDataRange().getValues();
    let maxOrd = 0;
    for (let i = 1; i < dbOrd.length; i++) {
      const m = (dbOrd[i][0] || '').toString().match(/ORD-\d{4}-(\d+)/);
      if (m) maxOrd = Math.max(maxOrd, parseInt(m[1]));
    }
    const idOrden = "ORD-" + anio + "-" + String(maxOrd + 1).padStart(4, "0");

    // Calcular financiero proporcional por grupo
    const totalGlobal = (payload.financiero || {}).total || 0;
    const subtotalGlobal = (payload.financiero || {}).subtotal || 0;
    const ivaGlobal = (payload.financiero || {}).iva || 0;
    const sumaVD = itemsVentaDirecta.reduce((s, it) => s + (it.total || 0), 0);
    const sumaPed = itemsProduccion.reduce((s, it) => s + (it.total || 0), 0);
    const totalItems = sumaVD + sumaPed || 1;

    function financieroProporcional(items) {
      const proporcion = items.reduce((s, it) => s + (it.total || 0), 0) / totalItems;
      return {
        subtotal: parseFloat((subtotalGlobal * proporcion).toFixed(2)),
        iva: parseFloat((ivaGlobal * proporcion).toFixed(2)),
        total: parseFloat((totalGlobal * proporcion).toFixed(2))
      };
    }

    // ── VENTA DIRECTA ────────────────────────────────────────────────────
    if (itemsVentaDirecta.length > 0) {
      let hojaVD = ss.getSheetByName("Ventas_Directas");
      if (!hojaVD) {
        hojaVD = ss.insertSheet("Ventas_Directas");
        hojaVD.appendRow(["ID_Venta","Fecha","Cliente","DUI_NIT","Tipo_Doc",
          "Items_Detalle","Subtotal","IVA","Total","Forma_Pago","Estado","Notas"]);
      }
      const dbVD = hojaVD.getDataRange().getValues();
      let maxVD = 0;
      for (let i = 1; i < dbVD.length; i++) {
        const m = (dbVD[i][0] || '').toString().match(/VD-\d{4}-(\d+)/);
        if (m) maxVD = Math.max(maxVD, parseInt(m[1]));
      }
      const idVD = "VD-" + anio + "-" + String(maxVD + 1).padStart(4, "0");
      const finVD = financieroProporcional(itemsVentaDirecta);

      hojaVD.appendRow([idVD, new Date(), cliente.nombre || "Consumidor Final",
        cliente.duiNit || "", cliente.tipoCliente || "Persona Natural",
        JSON.stringify(itemsVentaDirecta), finVD.subtotal, finVD.iva, finVD.total,
        formaPago, "Cobrada", (payload.notas || "") + (idOrden ? " [" + idOrden + "]" : "")]);

      registrarMovimientosSalidaPorVenta(itemsVentaDirecta, idVD, "Venta Directa");
      generarDocumentoFiscalServidor({ id: idVD, tipo: formaPago === 'Credito' ? 'Credito_Fiscal' : 'Consumidor_Final',
        tipoOperacion: 'Venta Directa', cliente, articulos: itemsVentaDirecta,
        financiero: finVD, formaPago, notas: payload.notas || '' });
      registrarIngresoServidor({ referencia: idVD + ' (' + idOrden + ')', tipo: 'Venta Directa',
        cliente: cliente.nombre || 'Consumidor Final', monto: finVD.total, formaPago });

      idVentaPrincipal = idVD;
      idsGenerados.push(idVD);
    }

    // ── PEDIDO A PRODUCCIÓN ──────────────────────────────────────────────
    if (itemsProduccion.length > 0) {
      let hojaPP = ss.getSheetByName("Pedidos_Produccion");
      if (!hojaPP) {
        hojaPP = ss.insertSheet("Pedidos_Produccion");
        hojaPP.appendRow(["ID_Pedido","Fecha","Cliente","DUI_NIT","Tipo_Doc",
          "Items_Detalle","Subtotal","IVA","Total","Estado","Notas"]);
      }
      const dbPP = hojaPP.getDataRange().getValues();
      let maxPP = 0;
      for (let i = 1; i < dbPP.length; i++) {
        const m = (dbPP[i][0] || '').toString().match(/PED-\d{4}-(\d+)/);
        if (m) maxPP = Math.max(maxPP, parseInt(m[1]));
      }
      const idPED = "PED-" + anio + "-" + String(maxPP + 1).padStart(4, "0");
      const finPED = financieroProporcional(itemsProduccion);

      hojaPP.appendRow([idPED, new Date(), cliente.nombre || "",
        cliente.duiNit || "", cliente.tipoCliente || "",
        JSON.stringify(itemsProduccion), finPED.subtotal, finPED.iva, finPED.total,
        "Pendiente de Producción", (payload.notas || "") + " [" + idOrden + "]"]);

      generarDocumentoFiscalServidor({ id: idPED, tipo: 'Orden_Produccion',
        tipoOperacion: 'Pedido a Producción', cliente, articulos: itemsProduccion,
        financiero: finPED, formaPago: 'Pendiente', notas: payload.notas || '' });
      registrarIngresoServidor({ referencia: idPED + ' (' + idOrden + ')', tipo: 'Cuenta por Cobrar',
        cliente: cliente.nombre || '', monto: finPED.total, formaPago: 'Pendiente de cobro' });

      if (!idVentaPrincipal) idVentaPrincipal = idPED;
      idsGenerados.push(idPED);
    }

    // Registrar la orden padre que agrupa VD y PED de esta operacion
    hojaOrdenes.appendRow([
      idOrden, new Date(), cliente.nombre || '', cliente.duiNit || '',
      JSON.stringify(articulos),
      (payload.financiero || {}).subtotal || 0,
      (payload.financiero || {}).iva || 0,
      (payload.financiero || {}).total || 0,
      formaPago, "Procesada", payload.notas || "",
      idsGenerados.join(', ')
    ]);

    const mensajeIds = idsGenerados.join(' + ');
    const tieneAmbos = itemsVentaDirecta.length > 0 && itemsProduccion.length > 0;
    const mensaje = tieneAmbos
      ? '¡Orden ' + idOrden + ': ' + idsGenerados[0] + ' (cobrada) y ' + idsGenerados[1] + ' enviado a Producción.'
      : '¡' + (itemsProduccion.length > 0 ? 'Pedido ' : 'Venta ') + mensajeIds + ' — Orden ' + idOrden + (itemsProduccion.length > 0 ? ' enviado a Producción!' : ' cobrada!');

    registrarAuditoria("Ventas", "Registrar", "Orden procesada: " + idOrden + " -> " + mensajeIds);
    return { exito: true, mensaje, id: idVentaPrincipal, idOrden, ids: idsGenerados };

  } catch (error) {
    return { exito: false, mensaje: "Error en venta: " + error.message };
  }
}


// ==========================================
//        MÓDULO: MOTOR DE CONFIGURACIÓN
// ==========================================

function obtenerConfiguracionIVA() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaConf = ss.getSheetByName("Configuracion");
    if (!hojaConf) return "SÍ";
    return hojaConf.getRange("B1").getValue().toString().toUpperCase() === "SÍ" ? "SÍ" : "NO";
  } catch (e) {
    return "SÍ";
  }
}

function verificarLicenciaProServidor() {
  return true;
}

/**
 * Compatible con el ModVentas viejo: devuelve productos disponibles para
 * VENTA del catalogo real del ERP (no la hoja Configuracion vieja).
 * Formato simplificado: { descripcion, precioBase, sku } por variante.
 */
function obtenerCatalogoProductos() {
  try {
    return obtenerCatalogoParaVentasServidor()
      .flatMap(p => (p.variantes || []).map(v => ({
        descripcion: p.nombre + (v.etiqueta ? ' - ' + v.etiqueta : ''),
        precioBase: v.pvp || p.pvp || 0,
        sku: v.sku
      })));
  } catch (e) {
    console.error("Error al obtener catálogo de productos: " + e.message);
    return [];
  }
}

/**
 * Devuelve el catalogo completo disponible para VENTA -- filtra por
 * disponibilidad de venta, incluye variantes con su PVP individual, y
 * aplana para que el buscador de Ventas pueda buscar por nombre, SKU
 * o etiqueta. Es la fuente de verdad del catalogo de Ventas.
 */
function obtenerCatalogoParaVentasServidor() {
  try {
    const productos = obtenerProductosServidor();
    const mapaDisponibilidad = obtenerMapaDisponibilidadServidor();
    const mapaCostos = obtenerMapaCostosServidor();

    return productos
      .filter(p => {
        if (p.tipo === 'Insumo') return false; // Insumos no se venden directamente
        const dispoProducto = mapaDisponibilidad["Producto:" + p.nombre];
        const dispoCategoria = mapaDisponibilidad["Categoria:" + p.categoria];
        const ventaProducto = dispoProducto ? dispoProducto.venta : true;
        const ventaCategoria = dispoCategoria ? dispoCategoria.venta : true;
        return ventaProducto && ventaCategoria;
      })
      .map(p => ({
        idProducto: p.idProducto,
        nombre: p.nombre,
        tipo: p.tipo,
        categoria: p.categoria,
        pvp: p.pvp || 0,
        variantes: (p.variantes || []).map(v => ({
          sku: v.sku,
          etiqueta: v.etiqueta || 'Estándar',
          pvp: v.pvp || p.pvp || mapaCostos[v.sku] || 0
        }))
      }));
  } catch (e) {
    console.error("Error al obtener catalogo para ventas: " + e.message);
    return [];
  }
}

/**
 * Devuelve las etiquetas/especificaciones disponibles desde el catalogo
 * real de variantes (no la hoja Configuracion vieja).
 */
function obtenerEtiquetasConfiguradas() {
  try {
    const productos = obtenerProductosServidor();
    const etiquetas = new Set();
    productos.forEach(p => {
      (p.variantes || []).forEach(v => {
        if (v.etiqueta && v.etiqueta.trim()) etiquetas.add(v.etiqueta.trim());
      });
    });
    return Array.from(etiquetas).sort();
  } catch (e) {
    console.error("Error al obtener etiquetas: " + e.message);
    return [];
  }
}



// =======================================================
//  PERFIL CORPORATIVO COMPLETO
// =======================================================
function guardarConfiguracionEmpresaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Config_Empresa");

    if (!sheet) {
      sheet = ss.insertSheet("Config_Empresa");
      sheet.appendRow([
        "Nombre Empresa", "NIT", "NRC", "Logo URL", "Giro Comercial", "Nombre ERP",
        "Dirección", "Teléfono", "Correo", "Régimen", "Clasificación", "CIIU",
        "Departamento", "Municipio", "Distrito", "Ambiente Hacienda", "Logo DTE",
        "Moneda Secundaria", "Tasa IVA"
      ]);
    }

    sheet.getRange(2, 1, 1, 19).setValues([[
      datos.nombreEmpresa,
      "'" + datos.nit,
      "'" + datos.nrc,
      datos.logoUrl,
      datos.giroComercial,
      datos.nombreErp,
      datos.direccion,
      datos.telefono,
      datos.correo,
      datos.regimen || "INFORMAL",
      datos.clasificacion || "Otros",
      datos.ciiu || "",
      datos.depto || "",
      datos.municipio || "",
      datos.distrito || "",
      datos.ambienteHacienda || "00",
      datos.logoDte || "",
      datos.monedaSecundaria || "NINGUNA",
      datos.tasaIva || 13
    ]]);

    return { exito: true };
  } catch (error) {
    return { exito: false, error: error.toString() };
  }
}

function obtenerConfiguracionEmpresaServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Config_Empresa");
    if (!sheet) return null;

    const datos = sheet.getRange(2, 1, 1, 19).getValues()[0];
    if (!datos[0] && !datos[5]) return null;

    return {
      nombreEmpresa: datos[0],
      nit: datos[1],
      nrc: datos[2],
      logoUrl: datos[3],
      giroComercial: datos[4],
      nombreErp: datos[5],
      direccion: datos[6],
      telefono: datos[7],
      correo: datos[8],
      regimen: datos[9],
      clasificacion: datos[10],
      ciiu: datos[11],
      depto: datos[12],
      municipio: datos[13],
      distrito: datos[14],
      ambienteHacienda: datos[15],
      logoDte: datos[16],
      monedaSecundaria: datos[17],
      tasaIva: datos[18]
    };
  } catch (error) {
    return null;
  }
}

// ==========================================
//    MÓDULO: MAESTRO DE PRODUCTOS E ÍTEMS
// ==========================================
//
// MODELO DE DATOS: Producto Base + Variantes
// -------------------------------------------
// Hoja "Productos" (producto base / "carpeta"):
//   1 ID_Producto | 2 Tipo | 3 Nombre | 4 Categoría | 5 Subcategoría |
//   6 Tributo DTE | 7 Unidad MH | 8 Flujos (JSON) | 9 Fecha Registro
//
// Hoja "Producto_Variantes" (cada variante vendible/comprable, con su propio stock):
//   1 ID_Producto (FK) | 2 SKU | 3 Etiqueta/Atributo (ej. "Azul") |
//   4 Costo | 5 Margen | 6 PVP | 7 Tags adicionales (JSON) | 8 Fecha Registro
//
// Un producto SIEMPRE tiene al menos 1 variante. Si el negocio no necesita
// distinguir variantes (ej. un servicio simple), la variante única puede
// usar la etiqueta "Estándar" y listo — el modelo no obliga a complicarse.

/**
 * Guarda un producto base junto con su primera variante (alta inicial).
 * @param {Object} datos { tipo, nombre, categoria, subcategoria, tributoDte, unidadDte, flujos,
 *                          variante: { sku, etiqueta, costo, margen, pvp, tags } }
 */
function guardarProductoServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaProd = ss.getSheetByName("Productos");
    let hojaVar = ss.getSheetByName("Producto_Variantes");

    if (!hojaProd) {
      hojaProd = ss.insertSheet("Productos");
      hojaProd.appendRow(["ID_Producto", "Tipo", "Nombre", "Categoría", "Subcategoría", "Tributo DTE", "Unidad MH", "Flujos (JSON)", "Fecha Registro"]);
    }
    if (!hojaVar) {
      hojaVar = ss.insertSheet("Producto_Variantes");
      hojaVar.appendRow(["ID_Producto", "SKU", "Etiqueta", "Costo", "Margen", "PVP", "Tags (JSON)", "Fecha Registro"]);
    }

    const idProducto = "PRD-" + new Date().getTime();

    hojaProd.appendRow([
      idProducto,
      datos.tipo,
      "'" + datos.nombre,
      "'" + datos.categoria,
      datos.subcategoria ? ("'" + datos.subcategoria) : "",
      datos.tributoDte || "1",
      datos.unidadDte,
      JSON.stringify(datos.flujos || {}),
      new Date()
    ]);

    return { exito: true, mensaje: "¡Producto guardado con éxito! Ahora agrégale al menos una variante.", idProducto: idProducto };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor de productos: " + error.message };
  }
}

/**
 * Agrega una nueva variante a un producto base ya existente.
 * @param {Object} datos { idProducto, sku, etiqueta, costo, margen, pvp, tags }
 */
function agregarVarianteServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaVar) {
      hojaVar = ss.insertSheet("Producto_Variantes");
      hojaVar.appendRow(["ID_Producto", "SKU", "Etiqueta", "Costo", "Margen", "PVP", "Tags (JSON)", "Fecha Registro"]);
    }

    const sku = (datos.sku || "").toString().trim();
    if (!sku) {
      return { exito: false, mensaje: "La variante necesita un SKU." };
    }
    if (skuYaExiste(sku)) {
      return { exito: false, mensaje: "⚠️ Ya existe una variante registrada con el SKU: " + sku };
    }

    hojaVar.appendRow([
      datos.idProducto,
      "'" + sku,
      "'" + (datos.etiqueta || "Estándar"),
      datos.costo || 0,
      datos.margen || 0,
      datos.pvp || 0,
      JSON.stringify(datos.tags || []),
      new Date()
    ]);

    return { exito: true, mensaje: "¡Variante añadida con éxito!", sku: sku };

  } catch (error) {
    return { exito: false, mensaje: "Error al añadir variante: " + error.message };
  }
}

/**
 * Actualiza los datos generales del producto base (no toca sus variantes).
 */
function actualizarProductoServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("Productos");
    if (!hojaProd) return { exito: false, mensaje: "No existe el catálogo de Productos todavía." };

    const db = hojaProd.getDataRange().getValues();
    let filaDestino = -1;

    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() === datos.idProducto.toString()) {
        filaDestino = i + 1;
        break;
      }
    }

    if (filaDestino === -1) {
      return { exito: false, mensaje: "No se encontró el producto original para actualizar." };
    }

    hojaProd.getRange(filaDestino, 2, 1, 7).setValues([[
      datos.tipo,
      "'" + datos.nombre,
      "'" + datos.categoria,
      datos.subcategoria ? ("'" + datos.subcategoria) : "",
      datos.tributoDte || "1",
      datos.unidadDte,
      JSON.stringify(datos.flujos || {})
    ]]);

    return { exito: true, mensaje: "¡Producto actualizado correctamente!", idProducto: datos.idProducto };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor al actualizar producto: " + error.message };
  }
}

/**
 * Actualiza una variante existente, localizada por su SKU original.
 */
function actualizarVarianteServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaVar) return { exito: false, mensaje: "No existen variantes registradas todavía." };

    const db = hojaVar.getDataRange().getValues();
    const skuBuscado = (datos.skuOriginal || datos.sku).toString().trim().toLowerCase();
    let filaDestino = -1;

    for (let i = 1; i < db.length; i++) {
      if (db[i][1].toString().trim().toLowerCase() === skuBuscado) {
        filaDestino = i + 1;
        break;
      }
    }

    if (filaDestino === -1) {
      return { exito: false, mensaje: "No se encontró la variante original para actualizar." };
    }

    const nuevoSku = datos.sku.toString().trim();
    if (nuevoSku.toLowerCase() !== skuBuscado && skuYaExiste(nuevoSku)) {
      return { exito: false, mensaje: "⚠️ Ya existe otra variante con ese nuevo SKU." };
    }

    hojaVar.getRange(filaDestino, 2, 1, 6).setValues([[
      "'" + nuevoSku,
      "'" + (datos.etiqueta || "Estándar"),
      datos.costo || 0,
      datos.margen || 0,
      datos.pvp || 0,
      JSON.stringify(datos.tags || [])
    ]]);

    return { exito: true, mensaje: "¡Variante actualizada correctamente!", sku: nuevoSku };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor al actualizar variante: " + error.message };
  }
}

/**
 * Elimina un producto base COMPLETO junto con todas sus variantes.
 * Se usa cuando se elimina desde la ficha principal del catálogo.
 */
function eliminarProductoServidor(idProducto) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("Productos");
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaProd) return { exito: false, mensaje: "No existe el catálogo de Productos." };

    const dbProd = hojaProd.getDataRange().getValues();
    let filaAEliminar = -1;
    for (let i = 1; i < dbProd.length; i++) {
      if (dbProd[i][0].toString() === idProducto.toString()) {
        filaAEliminar = i + 1;
        break;
      }
    }

    if (filaAEliminar !== -1) {
      hojaProd.deleteRow(filaAEliminar);
    }

    // Eliminar todas las variantes asociadas (de abajo hacia arriba para no desfasar índices)
    if (hojaVar) {
      const dbVar = hojaVar.getDataRange().getValues();
      for (let i = dbVar.length - 1; i >= 1; i--) {
        if (dbVar[i][0].toString() === idProducto.toString()) {
          hojaVar.deleteRow(i + 1);
        }
      }
    }

    return { exito: true, mensaje: "¡Producto y todas sus variantes fueron eliminados correctamente!" };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor al eliminar producto: " + error.message };
  }
}

/**
 * Elimina ÚNICAMENTE una variante puntual (no borra el producto base).
 * Bloquea si es la última variante restante del producto (un producto no puede quedar sin variantes).
 */
function eliminarVarianteServidor(sku) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaVar) return { exito: false, mensaje: "No existen variantes registradas." };

    const db = hojaVar.getDataRange().getValues();
    let filaAEliminar = -1;
    let idProductoDeEsta = null;

    for (let i = 1; i < db.length; i++) {
      if (db[i][1].toString().trim().toLowerCase() === sku.toString().trim().toLowerCase()) {
        filaAEliminar = i + 1;
        idProductoDeEsta = db[i][0].toString();
        break;
      }
    }

    if (filaAEliminar === -1) {
      return { exito: false, mensaje: "No se encontró la variante para eliminar." };
    }

    const totalVariantesDeEseProducto = db.slice(1).filter(fila => fila[0].toString() === idProductoDeEsta).length;
    if (totalVariantesDeEseProducto <= 1) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: es la única variante de este producto. Elimina el producto completo si ya no lo necesitas." };
    }

    hojaVar.deleteRow(filaAEliminar);
    return { exito: true, mensaje: "¡Variante eliminada correctamente!" };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor al eliminar variante: " + error.message };
  }
}

/**
 * Helper interno: revisa si un SKU ya existe en CUALQUIER variante del sistema.
 */
function skuYaExiste(sku) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaVar) return false;

    const db = hojaVar.getDataRange().getValues();
    const skuLower = sku.toString().trim().toLowerCase();

    for (let i = 1; i < db.length; i++) {
      if (db[i][1].toString().trim().toLowerCase() === skuLower) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Recupera el catálogo completo: cada producto base con su arreglo de variantes anidado.
 * Esta es la estructura que consume el frontend de ModConfig_Productos.
 */
function obtenerProductosServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("Productos");
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaProd) {
      Logger.log("obtenerProductosServidor: no existe la hoja Productos");
      return [];
    }

    const datosProd = hojaProd.getDataRange().getValues();
    Logger.log("obtenerProductosServidor: filas en Productos = " + datosProd.length);
    if (datosProd.length <= 1) return [];

    const datosVar = hojaVar ? hojaVar.getDataRange().getValues() : [];
    Logger.log("obtenerProductosServidor: filas en Producto_Variantes = " + datosVar.length);

    let productos = [];
    for (let i = 1; i < datosProd.length; i++) {
      const filaProd = datosProd[i];
      if (!filaProd || !filaProd[0]) continue;

      const idProducto = filaProd[0].toString();
      let flujosParsed = {};
      try { flujosParsed = JSON.parse(filaProd[7] || "{}"); } catch(e) { flujosParsed = {}; }

      const variantes = [];
      for (let j = 1; j < datosVar.length; j++) {
        const filaVar = datosVar[j];
        if (!filaVar || !filaVar[0]) continue;
        if (filaVar[0].toString() === idProducto) {
          let tagsParsed = [];
          try { tagsParsed = JSON.parse(filaVar[6] || "[]"); } catch(e) { tagsParsed = []; }
          variantes.push({
            sku: filaVar[1] ? filaVar[1].toString() : "",
            etiqueta: filaVar[2] ? filaVar[2].toString() : "",
            costo: Number(filaVar[3]) || 0,
            margen: Number(filaVar[4]) || 0,
            pvp: Number(filaVar[5]) || 0,
            tags: tagsParsed,
            fechaRegistro: filaVar[7] ? filaVar[7].toString() : ""
          });
        }
      }

      productos.push({
        idProducto: idProducto,
        tipo: filaProd[1] ? filaProd[1].toString() : "",
        nombre: filaProd[2] ? filaProd[2].toString() : "",
        categoria: filaProd[3] ? filaProd[3].toString() : "",
        subcategoria: filaProd[4] ? filaProd[4].toString() : "",
        tributoDte: filaProd[5] ? filaProd[5].toString() : "1",
        unidadDte: filaProd[6] ? filaProd[6].toString() : "05",
        flujos: flujosParsed,
        variantes: variantes,
        fechaRegistro: filaProd[8] ? filaProd[8].toString() : ""
      });
    }

    Logger.log("obtenerProductosServidor: productos armados = " + productos.length);
    return productos;

  } catch (e) {
    Logger.log("obtenerProductosServidor ERROR: " + e.message + " | " + e.stack);
    return [];
  }
}

// NOTA: guardarCategoriaServidor() fue reemplazada por crearCategoriaServidor(),
// que ahora soporta padre/subcategoría y bloquea duplicados nombre+padre.

/**
 * Crea una categoría o subcategoría desde ModConfig_Productos.
 * @param {Object} datosCategoria { nombre, padre } — padre = "NINGUNA" si es categoría raíz
 * Evita duplicados exactos (mismo nombre + mismo padre).
 * Devuelve la lista completa actualizada de categorías para repintar el selector.
 */
function crearCategoriaServidor(datosCategoria) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Categorias");

    if (!hoja) {
      hoja = ss.insertSheet("Categorias");
      hoja.appendRow(["Nombre Categoría", "Dependencia / Padre", "Fecha Creación"]);
    }

    const nombre = (datosCategoria && typeof datosCategoria === 'object') ? datosCategoria.nombre : datosCategoria;
    const padre = (datosCategoria && typeof datosCategoria === 'object' && datosCategoria.padre) ? datosCategoria.padre : "NINGUNA";

    const nombreLimpio = nombre.toString().trim();
    const padreLimpio = padre.toString().trim();

    if (!nombreLimpio) {
      return { exito: false, mensaje: "El nombre de la categoría no puede estar vacío.", categorias: obtenerCategoriasServidor() };
    }

    const db = hoja.getDataRange().getValues();

    for (let i = 1; i < db.length; i++) {
      const nombreFila = db[i][0].toString().trim().toLowerCase();
      const padreFila = (db[i][1] || "NINGUNA").toString().trim().toLowerCase();
      if (nombreFila === nombreLimpio.toLowerCase() && padreFila === padreLimpio.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Esa categoría ya existe bajo la misma dependencia.", categorias: obtenerCategoriasServidor() };
      }
    }

    hoja.appendRow(["'" + nombreLimpio, padreLimpio, new Date()]);
    return { exito: true, mensaje: "¡Categoría creada con éxito!", categorias: obtenerCategoriasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al crear categoría: " + error.message, categorias: [] };
  }
}

/**
 * Devuelve todas las categorías registradas, en formato { nombre, padre }
 */
function obtenerCategoriasServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Categorias");
    if (!hoja) return [];

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    let categorias = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") {
        categorias.push({
          nombre: datos[i][0].toString().trim(),
          padre: datos[i][1] ? datos[i][1].toString().trim() : "NINGUNA"
        });
      }
    }
    return categorias;
  } catch (e) {
    console.error("Error al obtener categorías: " + e.message);
    return [];
  }
}

/**
 * Edita el nombre de una categoría existente (localizada por nombre+padre originales).
 * Bloquea la edición si algún producto la está usando como categoría o subcategoría.
 */
function editarCategoriaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Categorias");
    if (!hoja) return { exito: false, mensaje: "No existe el almacén de categorías.", categorias: [] };

    const nombreOriginal = (datos.nombreOriginal || "").toString().trim();
    const padreOriginal = (datos.padreOriginal || "NINGUNA").toString().trim();
    const nombreNuevo = (datos.nombreNuevo || "").toString().trim();

    if (!nombreNuevo) {
      return { exito: false, mensaje: "El nuevo nombre no puede estar vacío.", categorias: obtenerCategoriasServidor() };
    }

    // Bloqueo: ¿algún producto usa esta categoría/subcategoría actualmente?
    if (categoriaEstaEnUso(nombreOriginal)) {
      return { exito: false, mensaje: "⚠️ No se puede editar: hay productos registrados usando esta categoría.", categorias: obtenerCategoriasServidor() };
    }

    const db = hoja.getDataRange().getValues();
    let filaDestino = -1;

    for (let i = 1; i < db.length; i++) {
      const nombreFila = db[i][0].toString().trim();
      const padreFila = (db[i][1] || "NINGUNA").toString().trim();
      if (nombreFila.toLowerCase() === nombreOriginal.toLowerCase() && padreFila.toLowerCase() === padreOriginal.toLowerCase()) {
        filaDestino = i + 1;
        break;
      }
    }

    if (filaDestino === -1) {
      return { exito: false, mensaje: "No se encontró la categoría original.", categorias: obtenerCategoriasServidor() };
    }

    // Evitar duplicado contra el nuevo nombre bajo el mismo padre
    for (let i = 1; i < db.length; i++) {
      if (i + 1 === filaDestino) continue;
      const nombreFila = db[i][0].toString().trim().toLowerCase();
      const padreFila = (db[i][1] || "NINGUNA").toString().trim().toLowerCase();
      if (nombreFila === nombreNuevo.toLowerCase() && padreFila === padreOriginal.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe otra categoría con ese nombre bajo la misma dependencia.", categorias: obtenerCategoriasServidor() };
      }
    }

    hoja.getRange(filaDestino, 1).setValue("'" + nombreNuevo);

    // Si esta categoría es padre de otras (subcategorías), actualizar su referencia de padre también
    for (let i = 1; i < db.length; i++) {
      const padreFila = (db[i][1] || "NINGUNA").toString().trim();
      if (padreFila.toLowerCase() === nombreOriginal.toLowerCase()) {
        hoja.getRange(i + 1, 2).setValue("'" + nombreNuevo);
      }
    }

    return { exito: true, mensaje: "¡Categoría actualizada con éxito!", categorias: obtenerCategoriasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al editar categoría: " + error.message, categorias: [] };
  }
}

/**
 * Elimina una categoría (localizada por nombre+padre).
 * Bloquea si: (a) tiene productos asociados, o (b) tiene subcategorías hijas.
 */
function eliminarCategoriaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Categorias");
    if (!hoja) return { exito: false, mensaje: "No existe el almacén de categorías.", categorias: [] };

    const nombre = (datos.nombre || "").toString().trim();
    const padre = (datos.padre || "NINGUNA").toString().trim();

    if (categoriaEstaEnUso(nombre)) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: hay productos registrados usando esta categoría.", categorias: obtenerCategoriasServidor() };
    }

    const db = hoja.getDataRange().getValues();

    // Bloqueo: si es categoría raíz, verificar que no tenga hijas
    const tieneHijas = db.slice(1).some(fila => (fila[1] || "NINGUNA").toString().trim().toLowerCase() === nombre.toLowerCase());
    if (tieneHijas) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: esta categoría tiene subcategorías dependientes. Elimínalas primero.", categorias: obtenerCategoriasServidor() };
    }

    let filaAEliminar = -1;
    for (let i = 1; i < db.length; i++) {
      const nombreFila = db[i][0].toString().trim();
      const padreFila = (db[i][1] || "NINGUNA").toString().trim();
      if (nombreFila.toLowerCase() === nombre.toLowerCase() && padreFila.toLowerCase() === padre.toLowerCase()) {
        filaAEliminar = i + 1;
        break;
      }
    }

    if (filaAEliminar === -1) {
      return { exito: false, mensaje: "No se encontró la categoría para eliminar.", categorias: obtenerCategoriasServidor() };
    }

    hoja.deleteRow(filaAEliminar);
    return { exito: true, mensaje: "¡Categoría eliminada con éxito!", categorias: obtenerCategoriasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al eliminar categoría: " + error.message, categorias: [] };
  }
}

/**
 * Helper interno: revisa si algún producto usa el nombre dado como Categoría o Subcategoría.
 */
function categoriaEstaEnUso(nombreCategoria) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("Productos");
    if (!hojaProd) return false;

    const datos = hojaProd.getDataRange().getValues();
    const nombreLower = nombreCategoria.toString().trim().toLowerCase();

    for (let i = 1; i < datos.length; i++) {
      const categoriaFila = (datos[i][3] || "").toString().trim().toLowerCase();
      const subcategoriaFila = (datos[i][4] || "").toString().trim().toLowerCase();
      if (categoriaFila === nombreLower || subcategoriaFila === nombreLower) {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error("Error al verificar uso de categoría: " + e.message);
    return false;
  }
}
function crearEtiquetaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Etiquetas_Operativas");

    if (!hoja) {
      hoja = ss.insertSheet("Etiquetas_Operativas");
      hoja.appendRow(["Nombre Etiqueta", "Grupo", "Fecha Creación"]);
    }

    const nombre = (datos && typeof datos === 'object') ? datos.nombre : datos;
    const grupo = (datos && typeof datos === 'object' && datos.grupo) ? datos.grupo.toString().trim() : "General";
    const nombreLimpio = nombre.toString().trim();

    if (!nombreLimpio) {
      return { exito: false, mensaje: "El nombre de la etiqueta no puede estar vacío.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    const db = hoja.getDataRange().getValues();

    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString().toLowerCase().trim() === nombreLimpio.toLowerCase() &&
          (db[i][1] || "General").toString().toLowerCase().trim() === grupo.toLowerCase()) {
        return { exito: true, mensaje: "La etiqueta ya existía en ese grupo.", etiquetas: obtenerEtiquetasOperativasServidor() };
      }
    }

    // Se antepone un apostrofe para forzar que Sheets guarde el nombre
    // SIEMPRE como texto literal -- sin esto, valores como "1/2" se
    // interpretan automaticamente como una fecha (1 de febrero) en vez de
    // conservarse como el texto que el usuario realmente escribio.
    hoja.appendRow(["'" + nombreLimpio, grupo, new Date()]);
    return { exito: true, mensaje: "Etiqueta creada con éxito.", etiquetas: obtenerEtiquetasOperativasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error en el servidor de etiquetas: " + error.message, etiquetas: [] };
  }
}

/**
 * Devuelve todas las etiquetas operativas globales, agrupadas: { nombre, grupo }
 */
function obtenerEtiquetasOperativasServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Etiquetas_Operativas");
    if (!hoja) return [];

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    let etiquetas = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") {
        etiquetas.push({
          nombre: datos[i][0].toString().trim(),
          grupo: datos[i][1] ? datos[i][1].toString().trim() : "General"
        });
      }
    }
    return etiquetas;
  } catch (e) {
    console.error("Error al obtener etiquetas operativas: " + e.message);
    return [];
  }
}

/**
 * Devuelve la lista única de grupos ya usados (para el selector "grupo existente").
 */
/**
 * Crea un grupo de etiquetas VACÍO (sin ninguna etiqueta dentro todavía).
 * Se usa cuando el usuario solo quiere crear la "carpeta" sin asignarle nada aún.
 */
function crearGrupoEtiquetaServidor(nombreGrupo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Grupos_Etiquetas");

    if (!hoja) {
      hoja = ss.insertSheet("Grupos_Etiquetas");
      hoja.appendRow(["Nombre Grupo", "Fecha Creación"]);
    }

    const nombreLimpio = nombreGrupo.toString().trim();
    if (!nombreLimpio) {
      return { exito: false, mensaje: "El nombre del grupo no puede estar vacío.", grupos: obtenerGruposEtiquetasServidor() };
    }

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString().toLowerCase().trim() === nombreLimpio.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe un grupo con ese nombre.", grupos: obtenerGruposEtiquetasServidor() };
      }
    }

    // También bloquea si ya existe como grupo derivado de alguna etiqueta
    const gruposExistentes = obtenerGruposEtiquetasServidor();
    if (gruposExistentes.map(g => g.toLowerCase()).includes(nombreLimpio.toLowerCase())) {
      return { exito: false, mensaje: "⚠️ Ya existe un grupo con ese nombre.", grupos: gruposExistentes };
    }

    hoja.appendRow(["'" + nombreLimpio, new Date()]);
    return { exito: true, mensaje: "¡Grupo creado con éxito!", grupos: obtenerGruposEtiquetasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al crear grupo: " + error.message, grupos: [] };
  }
}

/**
 * Devuelve TODOS los grupos: los creados explícitamente (vacíos o no) + los
 * que ya existen implícitamente porque alguna etiqueta los usa.
 */
function obtenerGruposEtiquetasServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaGrupos = ss.getSheetByName("Grupos_Etiquetas");

    let gruposExplicitos = [];
    if (hojaGrupos) {
      const datos = hojaGrupos.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        if (datos[i][0] && datos[i][0].toString().trim() !== "") {
          gruposExplicitos.push(datos[i][0].toString().trim());
        }
      }
    }

    const etiquetas = obtenerEtiquetasOperativasServidor();
    const gruposDeEtiquetas = etiquetas.map(e => e.grupo);

    const todos = [...new Set([...gruposExplicitos, ...gruposDeEtiquetas])];
    return todos.sort();
  } catch (e) {
    return [];
  }
}

/**
 * Edita el nombre de una etiqueta operativa existente.
 * Bloquea si algún producto la tiene asignada en su lista de tags.
 */
function editarEtiquetaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Etiquetas_Operativas");
    if (!hoja) return { exito: false, mensaje: "No existe el diccionario de etiquetas.", etiquetas: [] };

    const nombreOriginal = (datos.nombreOriginal || "").toString().trim();
    const grupoOriginal = (datos.grupoOriginal || "General").toString().trim();
    const nombreNuevo = (datos.nombreNuevo || "").toString().trim();

    if (!nombreNuevo) {
      return { exito: false, mensaje: "El nuevo nombre no puede estar vacío.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    if (etiquetaEstaEnUso(nombreOriginal)) {
      return { exito: false, mensaje: "⚠️ No se puede editar: hay productos que usan esta etiqueta.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    const db = hoja.getDataRange().getValues();
    let filaDestino = -1;

    for (let i = 1; i < db.length; i++) {
      const nombreFila = db[i][0].toString().trim();
      const grupoFila = (db[i][1] || "General").toString().trim();
      if (nombreFila.toLowerCase() === nombreOriginal.toLowerCase() && grupoFila.toLowerCase() === grupoOriginal.toLowerCase()) {
        filaDestino = i + 1;
        break;
      }
    }

    if (filaDestino === -1) {
      return { exito: false, mensaje: "No se encontró la etiqueta original.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    for (let i = 1; i < db.length; i++) {
      if (i + 1 === filaDestino) continue;
      const nombreFila = db[i][0].toString().trim().toLowerCase();
      const grupoFila = (db[i][1] || "General").toString().trim().toLowerCase();
      if (nombreFila === nombreNuevo.toLowerCase() && grupoFila === grupoOriginal.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe otra etiqueta con ese nombre en ese grupo.", etiquetas: obtenerEtiquetasOperativasServidor() };
      }
    }

    hoja.getRange(filaDestino, 1).setValue("'" + nombreNuevo);
    return { exito: true, mensaje: "¡Etiqueta actualizada con éxito!", etiquetas: obtenerEtiquetasOperativasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al editar etiqueta: " + error.message, etiquetas: [] };
  }
}

/**
 * Elimina una etiqueta operativa global (localizada por nombre+grupo). Bloquea si está en uso.
 */
function eliminarEtiquetaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Etiquetas_Operativas");
    if (!hoja) return { exito: false, mensaje: "No existe el diccionario de etiquetas.", etiquetas: [] };

    const nombre = (datos && typeof datos === 'object') ? datos.nombre : datos;
    const grupo = (datos && typeof datos === 'object' && datos.grupo) ? datos.grupo.toString().trim() : "General";
    const nombreLimpio = nombre.toString().trim();

    if (etiquetaEstaEnUso(nombreLimpio)) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: hay productos que usan esta etiqueta.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    const db = hoja.getDataRange().getValues();
    let filaAEliminar = -1;

    for (let i = 1; i < db.length; i++) {
      const nombreFila = db[i][0].toString().trim();
      const grupoFila = (db[i][1] || "General").toString().trim();
      if (nombreFila.toLowerCase() === nombreLimpio.toLowerCase() && grupoFila.toLowerCase() === grupo.toLowerCase()) {
        filaAEliminar = i + 1;
        break;
      }
    }

    if (filaAEliminar === -1) {
      return { exito: false, mensaje: "No se encontró la etiqueta para eliminar.", etiquetas: obtenerEtiquetasOperativasServidor() };
    }

    hoja.deleteRow(filaAEliminar);
    return { exito: true, mensaje: "¡Etiqueta eliminada con éxito!", etiquetas: obtenerEtiquetasOperativasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al eliminar etiqueta: " + error.message, etiquetas: [] };
  }
}

/**
 * Edita el nombre de un grupo de etiquetas. Renombra el grupo explícito
 * (hoja Grupos_Etiquetas) y también actualiza todas las etiquetas que
 * pertenecen a ese grupo para que apunten al nuevo nombre.
 */
function editarGrupoEtiquetaServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombreOriginal = (datos.nombreOriginal || "").toString().trim();
    const nombreNuevo = (datos.nombreNuevo || "").toString().trim();

    if (!nombreNuevo) {
      return { exito: false, mensaje: "El nuevo nombre no puede estar vacío.", grupos: obtenerGruposEtiquetasServidor() };
    }

    const gruposExistentes = obtenerGruposEtiquetasServidor();
    if (gruposExistentes.map(g => g.toLowerCase()).includes(nombreNuevo.toLowerCase()) && nombreNuevo.toLowerCase() !== nombreOriginal.toLowerCase()) {
      return { exito: false, mensaje: "⚠️ Ya existe otro grupo con ese nombre.", grupos: gruposExistentes };
    }

    // Renombrar en la hoja explícita de grupos, si existe ahí
    const hojaGrupos = ss.getSheetByName("Grupos_Etiquetas");
    if (hojaGrupos) {
      const dbGrupos = hojaGrupos.getDataRange().getValues();
      for (let i = 1; i < dbGrupos.length; i++) {
        if (dbGrupos[i][0].toString().trim().toLowerCase() === nombreOriginal.toLowerCase()) {
          hojaGrupos.getRange(i + 1, 1).setValue("'" + nombreNuevo);
          break;
        }
      }
    }

    // Actualizar todas las etiquetas que pertenecen a este grupo
    const hojaEtiquetas = ss.getSheetByName("Etiquetas_Operativas");
    if (hojaEtiquetas) {
      const dbEtiquetas = hojaEtiquetas.getDataRange().getValues();
      for (let i = 1; i < dbEtiquetas.length; i++) {
        const grupoFila = (dbEtiquetas[i][1] || "General").toString().trim();
        if (grupoFila.toLowerCase() === nombreOriginal.toLowerCase()) {
          hojaEtiquetas.getRange(i + 1, 2).setValue("'" + nombreNuevo);
        }
      }
    }

    return {
      exito: true,
      mensaje: "¡Grupo actualizado con éxito!",
      grupos: obtenerGruposEtiquetasServidor(),
      etiquetas: obtenerEtiquetasOperativasServidor()
    };

  } catch (error) {
    return { exito: false, mensaje: "Error al editar grupo: " + error.message, grupos: [] };
  }
}

/**
 * Elimina un grupo de etiquetas. Bloquea si el grupo todavía tiene
 * etiquetas asignadas (hay que eliminarlas o moverlas primero).
 */
function eliminarGrupoEtiquetaServidor(nombreGrupo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombreLimpio = nombreGrupo.toString().trim();

    const etiquetasDelGrupo = obtenerEtiquetasOperativasServidor().filter(e => e.grupo.toLowerCase() === nombreLimpio.toLowerCase());
    if (etiquetasDelGrupo.length > 0) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: este grupo todavía tiene etiquetas asignadas. Elimínalas primero.", grupos: obtenerGruposEtiquetasServidor() };
    }

    const hojaGrupos = ss.getSheetByName("Grupos_Etiquetas");
    if (hojaGrupos) {
      const db = hojaGrupos.getDataRange().getValues();
      for (let i = 1; i < db.length; i++) {
        if (db[i][0].toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
          hojaGrupos.deleteRow(i + 1);
          break;
        }
      }
    }

    return { exito: true, mensaje: "¡Grupo eliminado con éxito!", grupos: obtenerGruposEtiquetasServidor() };

  } catch (error) {
    return { exito: false, mensaje: "Error al eliminar grupo: " + error.message, grupos: [] };
  }
}

/**
 * Helper interno: revisa si alguna variante usa esta etiqueta, ya sea como
 * su atributo diferenciador (ej. "Azul") o dentro de sus tags adicionales.
 */
function etiquetaEstaEnUso(nombreEtiqueta) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    if (!hojaVar) return false;

    const datos = hojaVar.getDataRange().getValues();
    const nombreLower = nombreEtiqueta.toString().trim().toLowerCase();

    for (let i = 1; i < datos.length; i++) {
      const atributoFila = (datos[i][2] || "").toString().trim().toLowerCase();
      if (atributoFila === nombreLower) return true;

      let tagsParsed = [];
      try { tagsParsed = JSON.parse(datos[i][6] || "[]"); } catch(e) {}
      const tagsLower = tagsParsed.map(t => t.toString().trim().toLowerCase());
      if (tagsLower.includes(nombreLower)) return true;
    }
    return false;
  } catch (e) {
    console.error("Error al verificar uso de etiqueta: " + e.message);
    return false;
  }
}


// ==========================================
//        MODULO: INVENTARIOS
// ==========================================
//
// MODELO: kardex de movimientos (cada entrada/salida queda registrada).
// La cantidad disponible se CALCULA sumando movimientos, nunca se guarda
// como un numero fijo que se sobreescribe (eso pierde el historial).
//
// Hoja "Movimientos_Inventario":
//   1 ID_Movimiento | 2 Fecha | 3 SKU | 4 Tipo (Entrada/Salida) |
//   5 Origen (Compra Reventa/Ajuste Manual/Venta/Consumo Interno/Averia/Produccion) |
//   6 Cantidad | 7 Bodega | 8 Referencia | 9 Motivo/Nota | 10 Usuario
//
// IMPORTANTE: "Compra Consumo Interno" (oficina) NO genera fila en esta hoja
// porque no mueve stock vendible. Solo actualiza el costo promedio en
// "Costos_Inventario" y queda registrada en "Historial_Costos" para el reporte.
//
// Hoja "Bodegas": 1 Nombre | 2 Fecha Creacion
// Hoja "Limites_Inventario": 1 SKU | 2 Minimo | 3 Maximo | 4 Fecha
// Hoja "Costos_Inventario": 1 SKU | 2 CostoPromedio | 3 Fecha Actualizacion
// Hoja "Historial_Costos": 1 Fecha | 2 SKU | 3 Origen | 4 Cantidad | 5 CostoUnitario | 6 Nota
// Hoja "Disponibilidad_Productos": 1 Tipo(Producto/Categoria) | 2 Nombre | 3 DisponibleVenta | 4 DisponibleCompra

/**
 * Calcula la cantidad disponible de UNA variante sumando todos sus movimientos.
 * Si se indica bodega, solo suma movimientos de esa bodega; si no, suma TODAS (vista global).
 */
function calcularStockDeVariante(sku, bodega) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Movimientos_Inventario");
    if (!hoja) return 0;

    const datos = hoja.getDataRange().getValues();
    const skuLower = sku.toString().trim().toLowerCase();
    const bodegaFiltro = bodega ? bodega.toString().trim().toLowerCase() : null;
    let total = 0;

    for (let i = 1; i < datos.length; i++) {
      if ((datos[i][2] || "").toString().trim().toLowerCase() !== skuLower) continue;
      if (bodegaFiltro && (datos[i][6] || "").toString().trim().toLowerCase() !== bodegaFiltro) continue;

      const cantidad = parseFloat(datos[i][5]) || 0;
      total += (datos[i][3] === "Entrada") ? cantidad : -cantidad;
    }
    return total;
  } catch (e) {
    console.error("Error al calcular stock de variante: " + e.message);
    return 0;
  }
}

/**
 * Devuelve un mapa { sku: { total, porBodega: {bodega: cantidad} } } para TODAS las variantes,
 * recorriendo la hoja de movimientos una sola vez.
 */
function obtenerMapaStockServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Movimientos_Inventario");
    const mapa = {};
    if (!hoja) return mapa;

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const sku = (datos[i][2] || "").toString().trim();
      if (!sku) continue;
      const cantidad = parseFloat(datos[i][5]) || 0;
      const delta = (datos[i][3] === "Entrada") ? cantidad : -cantidad;
      const bodega = (datos[i][6] || "Sin Bodega").toString().trim();

      if (!mapa[sku]) mapa[sku] = { total: 0, porBodega: {} };
      mapa[sku].total += delta;
      mapa[sku].porBodega[bodega] = (mapa[sku].porBodega[bodega] || 0) + delta;
    }
    return mapa;
  } catch (e) {
    console.error("Error al obtener mapa de stock: " + e.message);
    return {};
  }
}

/**
 * Dado un SKU y el mapa de stock, devuelve la bodega con mayor cantidad
 * disponible. Se usa para descontar automaticamente de la bodega correcta
 * al registrar una venta directa, sin que el usuario tenga que elegirla.
 */
function resolverBodegaOptima(sku, mapaStock) {
  const info = mapaStock[sku];
  if (!info || !info.porBodega) return "Sin Bodega";
  let mejorBodega = "Sin Bodega";
  let mejorCantidad = -Infinity;
  Object.keys(info.porBodega).forEach(bodega => {
    const cant = info.porBodega[bodega] || 0;
    if (cant > mejorCantidad) {
      mejorCantidad = cant;
      mejorBodega = bodega;
    }
  });
  return mejorBodega;
}

/**
 * Antes de confirmar una venta o salida, valida que haya stock suficiente
 * para CADA item (vista global, sumando todas las bodegas). No registra
 * nada si falta alguno (todo o nada).
 */
function validarStockDisponibleParaVenta(articulos) {
  try {
    if (!articulos || articulos.length === 0) return { ok: true };

    const mapaStock = obtenerMapaStockServidor();
    const mapaVariantes = obtenerMapaVariantesConFlujoServidor();

    // Mapa descripcion → sku como fallback
    const mapaDesc = {};
    Object.keys(mapaVariantes).forEach(sku => {
      const d = (mapaVariantes[sku].descripcion || "").toLowerCase().trim();
      if (d) mapaDesc[d] = sku;
    });

    const avisosBodegas = []; // items que requieren despacho desde varias bodegas

    for (let i = 0; i < articulos.length; i++) {
      const item = articulos[i];
      let sku = (item.sku || "").toString().trim();
      if (!sku || !mapaVariantes[sku]) {
        const d = (item.desc || item.descripcion || "").toLowerCase().trim();
        sku = mapaDesc[d] || "";
      }
      if (!sku) continue;

      const infoVariante = mapaVariantes[sku];
      if (!infoVariante || !infoVariante.controlaStock) continue;

      const infoStock = mapaStock[sku] || { total: 0, porBodega: {} };
      const disponibleTotal = infoStock.total || 0;
      const solicitado = parseFloat(item.cant) || 0;

      // Stock insuficiente en total → bloquear
      if (solicitado > disponibleTotal) {
        return {
          ok: false,
          mensaje: "⚠️ Stock insuficiente para \"" + (infoVariante.descripcion || sku) + "\". " +
            "Disponible total: " + disponibleTotal + ", solicitado: " + solicitado + "."
        };
      }

      // Verificar si alguna bodega individual tiene todo lo solicitado
      const porBodega = infoStock.porBodega || {};
      const bodegaSuficiente = Object.keys(porBodega).find(b => (porBodega[b] || 0) >= solicitado);

      if (!bodegaSuficiente) {
        // Ninguna bodega tiene todo — necesita despacho parcial desde varias
        const detalleBodegas = Object.keys(porBodega)
          .filter(b => (porBodega[b] || 0) > 0)
          .map(b => b + ": " + porBodega[b])
          .join(", ");
        avisosBodegas.push({
          sku: sku,
          descripcion: infoVariante.descripcion || sku,
          solicitado: solicitado,
          disponibleTotal: disponibleTotal,
          detalleBodegas: detalleBodegas
        });
      }
    }

    if (avisosBodegas.length > 0) {
      // Hay stock suficiente en total pero se necesitan varias bodegas.
      // Devolvemos ok:true con un aviso para que el frontend informe al usuario
      // antes de confirmar -- el usuario decide si procede o espera.
      const detalleAviso = avisosBodegas.map(a =>
        "• " + a.descripcion + " (necesita: " + a.solicitado + ", stock por bodega: " + a.detalleBodegas + ")"
      ).join("\n");
      return {
        ok: true,
        aviso: true,
        mensajeAviso: "⚠️ El stock disponible está distribuido en varias bodegas. " +
          "El despacho se hará desde múltiples bodegas:\n\n" + detalleAviso +
          "\n\n¿Deseas continuar con el despacho parcial?"
      };
    }

    return { ok: true };
  } catch (e) {
    return { ok: false, mensaje: "Error al validar stock: " + e.message };
  }
}

/**
 * Registra los movimientos de SALIDA por una venta ya confirmada (no valida,
 * eso ya se hizo antes). Solo mueve las variantes con flow "Controla Stock".
 */
function registrarMovimientosSalidaPorVenta(articulos, idReferencia, origen) {
  try {
    if (!articulos || articulos.length === 0) return;

    const mapaVariantes = obtenerMapaVariantesConFlujoServidor();
    const mapaStock = obtenerMapaStockServidor(); // para resolver bodega optima
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = obtenerOCrearHojaMovimientos(ss);

    // Mapa adicional por descripcion normalizada -- fallback cuando no viene SKU
    const mapaDescripcion = {};
    Object.keys(mapaVariantes).forEach(sku => {
      const desc = (mapaVariantes[sku].descripcion || "").toLowerCase().trim();
      if (desc) mapaDescripcion[desc] = sku;
    });

    function resolverSku(item) {
      const skuDirecto = (item.sku || "").toString().trim();
      if (skuDirecto && mapaVariantes[skuDirecto]) return skuDirecto;
      const desc = (item.desc || item.descripcion || "").toLowerCase().trim();
      return mapaDescripcion[desc] || null;
    }

    const itemsConStock = [];
    articulos.forEach(item => {
      const sku = resolverSku(item);
      if (!sku) return;
      const info = mapaVariantes[sku];
      if (!info || !info.controlaStock) return;
      // Resuelve la bodega con mayor stock disponible para este SKU
      const bodegaOptima = resolverBodegaOptima(sku, mapaStock);
      itemsConStock.push({ ...item, skuResuelto: sku, bodegaResuelta: bodegaOptima });
    });

    if (itemsConStock.length === 0) {
      console.log("registrarMovimientosSalidaPorVenta: ningun item controla stock.");
      return;
    }

    // 1. Registra en Movimientos_Inventario con despacho inteligente por bodega
    // Si una bodega tiene todo → un solo movimiento. Si no → distribuye entre bodegas.
    const movimientosParaWHOUT = []; // para construir el JSON del recibo

    itemsConStock.forEach(item => {
      const porBodega = (mapaStock[item.skuResuelto] || {}).porBodega || {};
      const cantSolicitada = parseFloat(item.cant) || 0;

      // ¿Una sola bodega tiene todo?
      const bodegaSuficiente = Object.keys(porBodega).find(b => (porBodega[b] || 0) >= cantSolicitada);

      if (bodegaSuficiente) {
        // Despacho simple desde una bodega
        const idMov = "MOV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
        hojaMov.appendRow([
          idMov, new Date(), item.skuResuelto, "Salida", origen,
          cantSolicitada, bodegaSuficiente, idReferencia || "",
          "Venta: " + idReferencia, Session.getActiveUser().getEmail() || ""
        ]);
        movimientosParaWHOUT.push({ sku: item.skuResuelto, descripcion: item.desc || item.skuResuelto, cantidadRecibida: cantSolicitada, costoUnitario: parseFloat(item.prec || 0), bodega: bodegaSuficiente });
        item.bodegaResuelta = bodegaSuficiente;
      } else {
        // Despacho parcial desde múltiples bodegas — descuenta en orden de mayor a menor stock
        const bodegasOrdenadas = Object.keys(porBodega)
          .filter(b => (porBodega[b] || 0) > 0)
          .sort((a, b) => (porBodega[b] || 0) - (porBodega[a] || 0));

        let restante = cantSolicitada;
        bodegasOrdenadas.forEach(bodega => {
          if (restante <= 0) return;
          const disponibleEnBodega = Math.min(porBodega[bodega] || 0, restante);
          const idMov = "MOV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
          hojaMov.appendRow([
            idMov, new Date(), item.skuResuelto, "Salida", origen,
            disponibleEnBodega, bodega, idReferencia || "",
            "Venta parcial desde " + bodega + ": " + idReferencia,
            Session.getActiveUser().getEmail() || ""
          ]);
          movimientosParaWHOUT.push({ sku: item.skuResuelto, descripcion: item.desc || item.skuResuelto + " (" + bodega + ")", cantidadRecibida: disponibleEnBodega, costoUnitario: parseFloat(item.prec || 0), bodega: bodega });
          restante -= disponibleEnBodega;
        });
        item.bodegaResuelta = bodegasOrdenadas[0] || "Múltiples bodegas";
      }
    });

    // 2. Genera recibo WH-OUT en Recepciones con el mismo formato que WH-IN
    let hojaRec = ss.getSheetByName("Recepciones");
    if (!hojaRec) {
      hojaRec = ss.insertSheet("Recepciones");
      hojaRec.appendRow(["ID", "Fecha", "ID_Orden", "Items (JSON)", "Usuario", "Tipo Destino",
        "Recibido Por", "Entrega Directa A", "Firma Digital", "Tipo Movimiento", "Bodega"]);
    }

    const idSalida = generarCorrelativoRecepcionServidor(hojaRec, "WH-OUT");
    const bodegaPrincipal = itemsConStock[0].bodegaResuelta || "Sin Bodega";
    const itemsJson = JSON.stringify(movimientosParaWHOUT);

    hojaRec.appendRow([
      idSalida,
      new Date(),
      idReferencia || "",    // ID_Orden = referencia a la venta
      itemsJson,
      resolverNombreUsuarioPorCorreoServidor(Session.getActiveUser().getEmail()),
      "Entrega Directa",     // Tipo Destino
      "",                    // Recibido Por
      origen,                // Entrega Directa A = "Venta Directa"
      "",                    // Firma Digital
      "Salida",              // Tipo Movimiento
      bodegaPrincipal        // Bodega de origen real
    ]);

    console.log("WH-OUT " + idSalida + " | Bodega: " + bodegaPrincipal + " | Ref: " + idReferencia);

  } catch (e) {
    console.error("Error al registrar movimientos de salida: " + e.message);
  }
}

function obtenerOCrearHojaMovimientos(ss) {
  let hoja = ss.getSheetByName("Movimientos_Inventario");
  if (!hoja) {
    hoja = ss.insertSheet("Movimientos_Inventario");
    hoja.appendRow(["ID_Movimiento", "Fecha", "SKU", "Tipo", "Origen", "Cantidad", "Bodega", "Referencia", "Motivo", "Usuario"]);
  }
  return hoja;
}

/**
 * Helper interno: { sku: { controlaStock, esActivoFijo, descripcion, tipoProducto } }
 * cruzando Productos + Producto_Variantes. tipoProducto sirve para filtrar el
 * catalogo de Inventarios (solo "Producto" e "Insumo" pueden tener inventario).
 */
function obtenerMapaVariantesConFlujoServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("Productos");
    const hojaVar = ss.getSheetByName("Producto_Variantes");
    const mapa = {};
    if (!hojaProd || !hojaVar) return mapa;

    const datosProd = hojaProd.getDataRange().getValues();
    const infoPorId = {};
    for (let i = 1; i < datosProd.length; i++) {
      const idProducto = (datosProd[i][0] || "").toString();
      let flujos = {};
      try { flujos = JSON.parse(datosProd[i][7] || "{}"); } catch(e) {}
      infoPorId[idProducto] = {
        flujos: flujos,
        nombre: (datosProd[i][2] || "").toString(),
        tipoProducto: (datosProd[i][1] || "").toString()
      };
    }

    const datosVar = hojaVar.getDataRange().getValues();
    for (let i = 1; i < datosVar.length; i++) {
      const idProducto = (datosVar[i][0] || "").toString();
      const sku = (datosVar[i][1] || "").toString().trim();
      if (!sku) continue;
      const etiqueta = (datosVar[i][2] || "").toString();
      const infoProd = infoPorId[idProducto] || { flujos: {}, nombre: "", tipoProducto: "" };

      mapa[sku] = {
        controlaStock: !!infoProd.flujos.stock,
        esActivoFijo: !!infoProd.flujos.activo,
        descripcion: infoProd.nombre + (etiqueta ? (" - " + etiqueta) : ""),
        tipoProducto: infoProd.tipoProducto
      };
    }
    return mapa;
  } catch (e) {
    console.error("Error al obtener mapa de variantes: " + e.message);
    return {};
  }
}

// ==========================================
//        BODEGAS
// ==========================================

function crearBodegaServidor(nombre) {
  try {
    const nombreLimpio = nombre.toString().trim();
    if (!nombreLimpio) return { exito: false, mensaje: "El nombre de la bodega no puede estar vacío.", bodegas: obtenerBodegasServidor() };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Bodegas");
    if (!hoja) {
      hoja = ss.insertSheet("Bodegas");
      hoja.appendRow(["Nombre", "Fecha Creación"]);
    }

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe una bodega con ese nombre.", bodegas: obtenerBodegasServidor() };
      }
    }

    hoja.appendRow(["'" + nombreLimpio, new Date()]);
    return { exito: true, mensaje: "¡Bodega creada con éxito!", bodegas: obtenerBodegasServidor() };
  } catch (error) {
    return { exito: false, mensaje: "Error al crear bodega: " + error.message, bodegas: [] };
  }
}

function obtenerBodegasServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Bodegas");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    const bodegas = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") bodegas.push(datos[i][0].toString().trim());
    }
    return bodegas;
  } catch (e) {
    console.error("Error al obtener bodegas: " + e.message);
    return [];
  }
}

function eliminarBodegaServidor(nombre) {
  try {
    const nombreLimpio = nombre.toString().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hojaMov = ss.getSheetByName("Movimientos_Inventario");
    if (hojaMov) {
      const db = hojaMov.getDataRange().getValues();
      for (let i = 1; i < db.length; i++) {
        if ((db[i][6] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
          return { exito: false, mensaje: "⚠️ No se puede eliminar: esta bodega tiene movimientos registrados.", bodegas: obtenerBodegasServidor() };
        }
      }
    }

    const hoja = ss.getSheetByName("Bodegas");
    if (hoja) {
      const db = hoja.getDataRange().getValues();
      for (let i = 1; i < db.length; i++) {
        if ((db[i][0] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
          hoja.deleteRow(i + 1);
          break;
        }
      }
    }

    return { exito: true, mensaje: "¡Bodega eliminada con éxito!", bodegas: obtenerBodegasServidor() };
  } catch (error) {
    return { exito: false, mensaje: "Error al eliminar bodega: " + error.message, bodegas: [] };
  }
}

// ==========================================
//        ENTRADAS
// ==========================================

/**
 * ENTRADA POR COMPRA - REVENTA: el producto comprado SI sube el stock vendible.
 * Solucion temporal hasta que exista el modulo de Compras con orden+recepcion.
 * @param {Object} datos { sku, cantidad, bodega, proveedor, costoUnitario, referencia }
 */
function registrarEntradaCompraServidor(datos) {
  try {
    const sku = (datos.sku || "").toString().trim();
    const cantidad = parseFloat(datos.cantidad) || 0;
    const bodega = (datos.bodega || "Sin Bodega").toString().trim();

    if (!sku) return { exito: false, mensaje: "Selecciona una variante valida." };
    if (cantidad <= 0) return { exito: false, mensaje: "La cantidad debe ser mayor a cero." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerOCrearHojaMovimientos(ss);

    const idMov = "MOV-" + new Date().getTime();
    const costoUnitario = parseFloat(datos.costoUnitario) || 0;

    hoja.appendRow([
      idMov, new Date(), sku, "Entrada", "Compra Reventa",
      cantidad, bodega, datos.referencia || "",
      "Proveedor: " + (datos.proveedor || "No especificado") + (costoUnitario ? (" | Costo unitario: $" + costoUnitario) : ""),
      Session.getActiveUser().getEmail() || ""
    ]);

    if (costoUnitario > 0) actualizarCostoPromedioServidor(sku, cantidad, costoUnitario, "Compra Reventa");

    return { exito: true, mensaje: `¡Entrada de ${cantidad} unidades registrada con éxito!`, nuevoStock: calcularStockDeVariante(sku) };
  } catch (error) {
    return { exito: false, mensaje: "Error al registrar entrada de compra: " + error.message };
  }
}

/**
 * COMPRA DE CONSUMO INTERNO (oficina): NO mueve stock vendible. Solo
 * actualiza el costo promedio del insumo y deja registro en Historial_Costos
 * para que el reporte distinga este movimiento de una compra de reventa.
 * @param {Object} datos { sku, cantidad, costoUnitario, proveedor, nota }
 */
function registrarCompraConsumoInternoServidor(datos) {
  try {
    const sku = (datos.sku || "").toString().trim();
    const cantidad = parseFloat(datos.cantidad) || 0;
    const costoUnitario = parseFloat(datos.costoUnitario) || 0;

    if (!sku) return { exito: false, mensaje: "Selecciona una variante valida." };
    if (cantidad <= 0) return { exito: false, mensaje: "La cantidad debe ser mayor a cero." };

    actualizarCostoPromedioServidor(sku, cantidad, costoUnitario, "Compra Consumo Interno");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaHist = ss.getSheetByName("Historial_Costos");
    if (!hojaHist) {
      hojaHist = ss.insertSheet("Historial_Costos");
      hojaHist.appendRow(["Fecha", "SKU", "Origen", "Cantidad", "Costo Unitario", "Nota"]);
    }
    hojaHist.appendRow([
      new Date(), sku, "Compra Consumo Interno", cantidad, costoUnitario,
      "Proveedor: " + (datos.proveedor || "No especificado") + (datos.nota ? (" | " + datos.nota) : "")
    ]);

    return { exito: true, mensaje: "¡Compra de consumo interno registrada! Se actualizó el costo, sin afectar el stock disponible para venta." };
  } catch (error) {
    return { exito: false, mensaje: "Error al registrar compra de consumo interno: " + error.message };
  }
}

/**
 * AJUSTE MANUAL: carga inicial o correccion de inventario, sin proveedor/costo.
 * @param {Object} datos { sku, cantidad, bodega, motivo } — cantidad puede ser negativa
 */
function registrarAjusteManualServidor(datos) {
  try {
    const sku = (datos.sku || "").toString().trim();
    const cantidad = parseFloat(datos.cantidad);
    const bodega = (datos.bodega || "Sin Bodega").toString().trim();

    if (!sku) return { exito: false, mensaje: "Selecciona una variante valida." };
    if (isNaN(cantidad) || cantidad === 0) return { exito: false, mensaje: "Ingresa una cantidad distinta de cero." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerOCrearHojaMovimientos(ss);

    const idMov = "MOV-" + new Date().getTime();
    const esEntrada = cantidad > 0;

    hoja.appendRow([
      idMov, new Date(), sku, esEntrada ? "Entrada" : "Salida", "Ajuste Manual",
      Math.abs(cantidad), bodega, "", datos.motivo || "Ajuste de inventario",
      Session.getActiveUser().getEmail() || ""
    ]);

    return { exito: true, mensaje: "¡Ajuste de inventario registrado con éxito!", nuevoStock: calcularStockDeVariante(sku) };
  } catch (error) {
    return { exito: false, mensaje: "Error al registrar ajuste manual: " + error.message };
  }
}

// ==========================================
//        SALIDAS
// ==========================================

/**
 * SALIDA POR AVERIA O RAZON JUSTIFICABLE: la unica salida manual permitida
 * desde Inventarios. Las salidas por Venta y Produccion las registran esos
 * modulos automaticamente (registrarMovimientosSalidaPorVenta, y a futuro
 * el equivalente de Produccion).
 * @param {Object} datos { sku, cantidad, bodega, motivo, responsable }
 */
/**
 * Registra una SALIDA de bodega -- ya sea por Averia/Ajuste (perdida,
 * producto danado/vencido) o por Entrega a Produccion/Area (insumos que
 * salen de bodega hacia otra area interna, ej. Produccion). Ambos tipos
 * generan un recibo WH-OUT en la hoja "Recepciones" para trazabilidad,
 * igual que las entradas generan WH-IN.
 * @param {Object} datos { sku, cantidad, bodega, tipoSalida, motivo, responsable, areaDestino, firmaDigital }
 *   tipoSalida: "Averia" | "Entrega a Produccion"
 */
function registrarSalidaAveriaServidor(datos) {
  return ejecutarConLock(function() {
    const sku = (datos.sku || "").toString().trim();
    const cantidad = parseFloat(datos.cantidad) || 0;
    const bodega = (datos.bodega || "Sin Bodega").toString().trim();
    const tipoSalida = (datos.tipoSalida === "Entrega a Producción") ? "Entrega a Producción" : "Avería";
    const motivo = (datos.motivo || "").toString().trim();
    const responsable = (datos.responsable || "").toString().trim();
    const areaDestino = (datos.areaDestino || "").toString().trim();
    const firmaDigital = (datos.firmaDigital || "").toString();

    if (!sku) return { exito: false, mensaje: "Selecciona una variante válida." };
    if (cantidad <= 0) return { exito: false, mensaje: "La cantidad debe ser mayor a cero." };
    if (tipoSalida === "Avería" && !motivo) return { exito: false, mensaje: "Indica el motivo de la avería." };
    if (tipoSalida === "Entrega a Producción" && !areaDestino) return { exito: false, mensaje: "Indica el área/departamento que recibe." };

    const disponible = calcularStockDeVariante(sku, datos.bodega || null);
    if (cantidad > disponible) {
      return { exito: false, mensaje: `⚠️ Stock insuficiente. Disponible: ${disponible}, solicitado: ${cantidad}.` };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = obtenerOCrearHojaMovimientos(ss);

    const idMov = "MOV-" + new Date().getTime();
    const notaMovimiento = tipoSalida === "Avería"
      ? (motivo + (responsable ? (" | Responsable: " + responsable) : ""))
      : ("Entrega a: " + areaDestino + (motivo ? (" | " + motivo) : ""));

    hojaMov.appendRow([
      idMov, new Date(), sku, "Salida", tipoSalida === "Avería" ? "Avería / Ajuste" : "Entrega a Producción",
      cantidad, bodega, "", notaMovimiento,
      Session.getActiveUser().getEmail() || ""
    ]);

    // Genera el recibo WH-OUT correspondiente, igual al patron de WH-IN en
    // recepciones de mercaderia -- mismo lugar de trazabilidad para ambos
    // tipos de movimiento (entradas y salidas de bodega).
    let hojaRecepciones = ss.getSheetByName("Recepciones");
    if (!hojaRecepciones) {
      hojaRecepciones = ss.insertSheet("Recepciones");
      hojaRecepciones.appendRow(["ID", "Fecha", "ID_Orden", "Items (JSON)", "Usuario", "Tipo Destino", "Recibido Por", "Entrega Directa A", "Firma Digital", "Tipo Movimiento", "Bodega"]);
    }
    const idSalida = generarCorrelativoRecepcionServidor(hojaRecepciones, "WH-OUT");

    hojaRecepciones.appendRow([
      idSalida, new Date(), "", JSON.stringify([{ sku: sku, cantidadRecibida: cantidad, descripcion: sku }]),
      resolverNombreUsuarioPorCorreoServidor(Session.getActiveUser().getEmail()),
      tipoSalida, // se reutiliza "Tipo Destino" para guardar si fue Averia o Entrega a Produccion
      tipoSalida === "Entrega a Producción" ? "" : responsable,
      tipoSalida === "Entrega a Producción" ? areaDestino : "",
      firmaDigital, "Salida", bodega
    ]);

    return { exito: true, mensaje: "¡Salida registrada con éxito! Recibo: " + idSalida, nuevoStock: calcularStockDeVariante(sku), idSalida: idSalida };
  });
}

// ==========================================
//        COSTEO (PROMEDIO PONDERADO)
// ==========================================

/**
 * Actualiza el costo promedio ponderado de una variante con una nueva entrada.
 * Formula: ((stockActual * costoActual) + (cantidadNueva * costoNuevo)) / (stockActual + cantidadNueva)
 * Si stockActual es 0, el costo nuevo simplemente se adopta.
 */
function actualizarCostoPromedioServidor(sku, cantidadNueva, costoNuevo, origen) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Costos_Inventario");
    if (!hoja) {
      hoja = ss.insertSheet("Costos_Inventario");
      hoja.appendRow(["SKU", "Costo Promedio", "Fecha Actualización"]);
    }

    const stockActual = calcularStockDeVariante(sku);
    const db = hoja.getDataRange().getValues();
    let filaDestino = -1;
    let costoActual = 0;

    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === sku.toLowerCase()) {
        filaDestino = i + 1;
        costoActual = parseFloat(db[i][1]) || 0;
        break;
      }
    }

    let costoPromedioNuevo;
    const baseStock = Math.max(stockActual, 0);
    if (baseStock + cantidadNueva > 0) {
      costoPromedioNuevo = ((baseStock * costoActual) + (cantidadNueva * costoNuevo)) / (baseStock + cantidadNueva);
    } else {
      costoPromedioNuevo = costoNuevo;
    }

    if (filaDestino === -1) {
      hoja.appendRow([sku, costoPromedioNuevo, new Date()]);
    } else {
      hoja.getRange(filaDestino, 2, 1, 2).setValues([[costoPromedioNuevo, new Date()]]);
    }

    return costoPromedioNuevo;
  } catch (e) {
    console.error("Error al actualizar costo promedio: " + e.message);
    return 0;
  }
}

function obtenerMapaCostosServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Costos_Inventario");
    const mapa = {};
    if (!hoja) return mapa;
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const sku = (datos[i][0] || "").toString().trim();
      if (sku) mapa[sku] = parseFloat(datos[i][1]) || 0;
    }
    return mapa;
  } catch (e) {
    return {};
  }
}

// ==========================================
//        LIMITES MIN/MAX (PRO)
// ==========================================

/**
 * Devuelve un mapa { "SKU|Bodega": {minimo, maximo} }. La clave "Global" se usa
 * para el limite general (cuando no se filtra por bodega especifica).
 */
function obtenerMapaLimitesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Limites_Inventario");
    const mapa = {};
    if (!hoja) return mapa;
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const sku = (datos[i][0] || "").toString().trim();
      if (!sku) continue;
      const bodega = (datos[i][1] || "Global").toString().trim();
      // Blindaje: si la columna "Bodega" es puramente numerica, es una fila
      // con formato viejo/corrupto (de antes de que existiera esta columna)
      // -- se ignora para que no genere una clave de mapa sin sentido.
      if (!isNaN(bodega) && bodega !== "") continue;
      mapa[sku + "|" + bodega] = { minimo: parseFloat(datos[i][2]) || 0, maximo: parseFloat(datos[i][3]) || 0 };
    }
    return mapa;
  } catch (e) {
    console.error("Error al obtener mapa de limites: " + e.message);
    return {};
  }
}

/**
 * Devuelve los limites de UNA variante desglosados por bodega: { "Global": {minimo,maximo}, "BodegaA": {...} }
 * Se usa en el modal de Limites para poder cambiar de bodega sin recargar todo el inventario.
 */
function obtenerLimitesDeVarianteServidor(sku) {
  try {
    const mapaCompleto = obtenerMapaLimitesServidor();
    const skuLimpio = sku.toString().trim();
    const resultado = { "Global": { minimo: 0, maximo: 0 } };

    Object.keys(mapaCompleto).forEach(clave => {
      const partes = clave.split("|");
      if (partes[0] === skuLimpio) {
        resultado[partes[1]] = mapaCompleto[clave];
      }
    });

    return resultado;
  } catch (e) {
    console.error("Error al obtener limites de variante: " + e.message);
    return { "Global": { minimo: 0, maximo: 0 } };
  }
}

/**
 * Guarda minimo/maximo para una variante en una bodega especifica (o "Global"
 * si no se indica bodega, que es el limite que se usa en la Vista Global).
 * @param {Object} datos { sku, bodega, minimo, maximo }
 */
function guardarLimitesInventarioServidor(datos) {
  try {
    const sku = (datos.sku || "").toString().trim();
    const bodega = (datos.bodega || "Global").toString().trim();
    const minimo = parseFloat(datos.minimo) || 0;
    const maximo = parseFloat(datos.maximo) || 0;

    if (!sku) return { exito: false, mensaje: "Selecciona una variante valida." };
    if (minimo < 0 || maximo < 0) return { exito: false, mensaje: "Los valores no pueden ser negativos." };
    if (maximo > 0 && maximo < minimo) return { exito: false, mensaje: "El maximo no puede ser menor que el minimo." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Limites_Inventario");
    if (!hoja) {
      hoja = ss.insertSheet("Limites_Inventario");
      hoja.appendRow(["SKU", "Bodega", "Minimo", "Maximo", "Fecha Actualizacion"]);
    }

    const db = hoja.getDataRange().getValues();
    let filaDestino = -1;
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === sku.toLowerCase() &&
          (db[i][1] || "Global").toString().trim().toLowerCase() === bodega.toLowerCase()) {
        filaDestino = i + 1;
        break;
      }
    }

    if (filaDestino === -1) {
      hoja.appendRow([sku, bodega, minimo, maximo, new Date()]);
    } else {
      hoja.getRange(filaDestino, 3, 1, 3).setValues([[minimo, maximo, new Date()]]);
    }

    return { exito: true, mensaje: "¡Límites de inventario guardados con éxito!" };
  } catch (error) {
    return { exito: false, mensaje: "Error al guardar límites: " + error.message };
  }
}

// ==========================================
//        HABILITAR / DESHABILITAR (Venta / Compra)
// ==========================================

/**
 * Devuelve el mapa de disponibilidad: { "Producto:NombreProducto": {venta,compra}, "Categoria:NombreCat": {venta,compra} }
 * Si una combinacion no existe en la hoja, se asume DISPONIBLE por defecto (true/true).
 */
/**
 * Catalogo completo para el modulo de Requisiciones: a diferencia de
 * obtenerInventarioAgrupadoServidor (que solo trae Producto/Insumo con
 * control de stock), aqui se incluyen TODOS los tipos -- Producto, Insumo,
 * Servicio y Produccion -- ya que un solicitante puede necesitar pedir
 * cualquiera de ellos, no solo lo que vive en bodega.
 */
function obtenerCatalogoParaRequisicionesServidor() {
  try {
    const productos = obtenerProductosServidor();
    const mapaDisponibilidad = obtenerMapaDisponibilidadServidor();

    return productos
      .filter(p => {
        const dispoProducto = mapaDisponibilidad["Producto:" + p.nombre];
        const dispoCategoria = mapaDisponibilidad["Categoria:" + p.categoria];
        const compraProducto = dispoProducto ? dispoProducto.compra : true;
        const compraCategoria = dispoCategoria ? dispoCategoria.compra : true;
        return compraProducto && compraCategoria;
      })
      .map(p => ({
        idProducto: p.idProducto,
        nombre: p.nombre,
        tipo: p.tipo,
        categoria: p.categoria,
        variantes: (p.variantes || []).map(v => ({ sku: v.sku, etiqueta: v.etiqueta }))
      }));
  } catch (e) {
    console.error("Error al obtener catalogo para requisiciones: " + e.message);
    return [];
  }
}

function obtenerMapaDisponibilidadServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Disponibilidad_Productos");
    const mapa = {};
    if (!hoja) return mapa;
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const tipo = (datos[i][0] || "").toString().trim();
      const nombre = (datos[i][1] || "").toString().trim();
      if (!tipo || !nombre) continue;
      mapa[tipo + ":" + nombre] = {
        venta: datos[i][2] !== false && datos[i][2] !== "FALSE",
        compra: datos[i][3] !== false && datos[i][3] !== "FALSE"
      };
    }
    return mapa;
  } catch (e) {
    console.error("Error al obtener mapa de disponibilidad: " + e.message);
    return {};
  }
}

/**
 * Habilita/deshabilita un Producto o Categoria para Venta y/o Compra.
 * @param {Object} datos { tipo: "Producto"|"Categoria", nombre, disponibleVenta, disponibleCompra }
 */
function guardarDisponibilidadServidor(datos) {
  try {
    const tipo = (datos.tipo || "").toString().trim();
    const nombre = (datos.nombre || "").toString().trim();
    if (!tipo || !nombre) return { exito: false, mensaje: "Datos incompletos." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Disponibilidad_Productos");
    if (!hoja) {
      hoja = ss.insertSheet("Disponibilidad_Productos");
      hoja.appendRow(["Tipo", "Nombre", "Disponible Venta", "Disponible Compra"]);
    }

    const db = hoja.getDataRange().getValues();
    let filaDestino = -1;
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim() === tipo && (db[i][1] || "").toString().trim() === nombre) {
        filaDestino = i + 1;
        break;
      }
    }

    const disponibleVenta = !!datos.disponibleVenta;
    const disponibleCompra = !!datos.disponibleCompra;

    if (filaDestino === -1) {
      hoja.appendRow([tipo, nombre, disponibleVenta, disponibleCompra]);
    } else {
      hoja.getRange(filaDestino, 3, 1, 2).setValues([[disponibleVenta, disponibleCompra]]);
    }

    return { exito: true, mensaje: "¡Disponibilidad actualizada con éxito!" };
  } catch (error) {
    return { exito: false, mensaje: "Error al guardar disponibilidad: " + error.message };
  }
}

// ==========================================
//        VISTA AGRUPADA PRINCIPAL
// ==========================================

/**
 * Catalogo agrupado para el modulo de Inventarios: solo productos tipo
 * "Producto" o "Insumo" con flujo "Controla Stock" activo. Incluye stock
 * total + por bodega, limites, alerta, sugerencia, costo promedio, y
 * disponibilidad de venta/compra.
 * @param {String} bodegaFiltro Opcional: si se indica, el stock mostrado es solo de esa bodega.
 */
function obtenerInventarioAgrupadoServidor(bodegaFiltro) {
  try {
    const productos = obtenerProductosServidor();
    const mapaStock = obtenerMapaStockServidor();
    const mapaLimites = obtenerMapaLimitesServidor();
    const mapaCostos = obtenerMapaCostosServidor();
    const mapaDisponibilidad = obtenerMapaDisponibilidadServidor();

    return productos
      .filter(p => (p.flujos || {}).stock && (p.tipo === "Producto" || p.tipo === "Insumo"))
      .map(p => {
        const dispoProducto = mapaDisponibilidad["Producto:" + p.nombre] || { venta: true, compra: true };
        const dispoCategoria = mapaDisponibilidad["Categoria:" + p.categoria] || { venta: true, compra: true };

        let totalProducto = 0;
        let totalMinimo = 0;
        let totalMaximo = 0;
        let alertaProducto = "normal";

        const variantesConStock = (p.variantes || []).map(v => {
          const infoStock = mapaStock[v.sku] || { total: 0, porBodega: {} };
          const stockVariante = bodegaFiltro ? (infoStock.porBodega[bodegaFiltro] || 0) : infoStock.total;

          let alertaVariante = "normal";
          let sugerenciaCompra = 0;
          let bodegasEnAlerta = [];
          let limites = { minimo: 0, maximo: 0 };

          if (bodegaFiltro) {
            // Vista de una bodega especifica: se evalua solo el limite de esa bodega.
            limites = mapaLimites[v.sku + "|" + bodegaFiltro] || { minimo: 0, maximo: 0 };
            if (limites.minimo > 0 || limites.maximo > 0) {
              if (stockVariante <= 0) alertaVariante = "agotado";
              else if (stockVariante <= limites.minimo) alertaVariante = "bajo";
              if (alertaVariante !== "normal" && limites.maximo > stockVariante) {
                sugerenciaCompra = limites.maximo - stockVariante;
              }
            }
          } else {
            // Vista Global: revisa el limite "Global" (si existe) Y el limite
            // de CADA bodega individual con stock registrado -- si cualquiera
            // esta agotada o por debajo de su minimo, la variante se marca en
            // alerta, indicando en que bodega(s) especificamente, para no
            // tener que ir filtrando bodega por bodega para descubrirlo.
            limites = mapaLimites[v.sku + "|Global"] || { minimo: 0, maximo: 0 };
            if (limites.minimo > 0 || limites.maximo > 0) {
              if (stockVariante <= 0) alertaVariante = "agotado";
              else if (stockVariante <= limites.minimo) alertaVariante = "bajo";
              if (alertaVariante !== "normal" && limites.maximo > stockVariante) {
                sugerenciaCompra = limites.maximo - stockVariante;
              }
            }

            // Se revisan las bodegas que tienen un LIMITE definido para este
            // SKU (no solo las que ya tuvieron algun movimiento registrado)
            // -- una bodega con stock en cero porque nunca recibio nada
            // jamas aparece en infoStock.porBodega, pero igual debe activar
            // la alerta si tiene un minimo definido.
            const nombresBodegasConLimite = Object.keys(mapaLimites)
              .filter(clave => clave.startsWith(v.sku + "|") && clave !== (v.sku + "|Global"))
              .map(clave => clave.substring((v.sku + "|").length));

            nombresBodegasConLimite.forEach(nombreBodega => {
              const limitesBodega = mapaLimites[v.sku + "|" + nombreBodega];
              if (!limitesBodega || (limitesBodega.minimo <= 0 && limitesBodega.maximo <= 0)) return;

              const stockEnEsaBodega = (infoStock.porBodega || {})[nombreBodega] || 0;
              let alertaEnEsaBodega = "normal";
              if (stockEnEsaBodega <= 0) alertaEnEsaBodega = "agotado";
              else if (stockEnEsaBodega <= limitesBodega.minimo) alertaEnEsaBodega = "bajo";

              if (alertaEnEsaBodega !== "normal") {
                bodegasEnAlerta.push({ bodega: nombreBodega, alerta: alertaEnEsaBodega, stock: stockEnEsaBodega, minimo: limitesBodega.minimo, maximo: limitesBodega.maximo });
                if (alertaEnEsaBodega === "agotado") alertaVariante = "agotado";
                else if (alertaVariante !== "agotado") alertaVariante = "bajo";

                if (limitesBodega.maximo > stockEnEsaBodega) {
                  const sugerenciaDeEstaBodega = limitesBodega.maximo - stockEnEsaBodega;
                  if (sugerenciaDeEstaBodega > sugerenciaCompra) sugerenciaCompra = sugerenciaDeEstaBodega;
                }
              }
            });
          }

          totalProducto += stockVariante;
          totalMinimo += limites.minimo;
          totalMaximo += limites.maximo;

          if (alertaVariante === "agotado") alertaProducto = "agotado";
          else if (alertaVariante === "bajo" && alertaProducto !== "agotado") alertaProducto = "bajo";

          return Object.assign({}, v, {
            stock: stockVariante,
            stockPorBodega: infoStock.porBodega,
            minimo: limites.minimo,
            maximo: limites.maximo,
            alerta: alertaVariante,
            bodegasEnAlerta: bodegasEnAlerta,
            sugerenciaCompra: sugerenciaCompra,
            costoPromedio: mapaCostos[v.sku] || v.costo || 0
          });
        });

        return Object.assign({}, p, {
          variantes: variantesConStock,
          stockTotal: totalProducto,
          minimoTotal: totalMinimo,
          maximoTotal: totalMaximo,
          alerta: alertaProducto,
          disponibleVenta: dispoProducto.venta && dispoCategoria.venta,
          disponibleCompra: dispoProducto.compra && dispoCategoria.compra
        });
      });
  } catch (e) {
    console.error("Error al obtener inventario agrupado: " + e.message);
    return [];
  }
}

/**
 * Exportacion a Excel. Estandar: columnas basicas. PRO: agrega bodega,
 * costeo, min/max, alerta, sugerencia, y valor en stock.
 */
function obtenerDatosExportacionInventarioServidor(esPro) {
  try {
    const inventario = obtenerInventarioAgrupadoServidor();
    const mapaLimites = obtenerMapaLimitesServidor();
    const filas = [];

    const encabezados = esPro
      ? ["Producto", "Categoría", "SKU", "Etiqueta", "Bodega", "Stock en Bodega", "Método Costeo", "Costo Promedio", "Mínimo", "Máximo", "Alerta", "Sugerencia de Compra", "PVP", "Valor en Stock (Costo)", "Disponible Venta", "Disponible Compra"]
      : ["Producto", "Categoría", "SKU", "Etiqueta", "Bodega", "Stock en Bodega"];

    filas.push(encabezados);

    function calcularAlerta(stock, limites) {
      if (limites.minimo <= 0 && limites.maximo <= 0) return "Normal";
      if (stock <= 0) return "AGOTADO";
      if (stock <= limites.minimo) return "STOCK BAJO";
      return "Normal";
    }

    inventario.forEach(p => {
      (p.variantes || []).forEach(v => {
        const bodegas = Object.entries(v.stockPorBodega || {});

        if (bodegas.length === 0) {
          // Variante sin movimientos registrados todavia
          if (esPro) {
            const limites = mapaLimites[v.sku + "|Global"] || { minimo: 0, maximo: 0 };
            filas.push([
              p.nombre, p.categoria || "", v.sku, v.etiqueta || "Estándar", "Sin Bodega", 0,
              "Promedio Ponderado", v.costoPromedio || 0, limites.minimo, limites.maximo,
              calcularAlerta(0, limites), 0, v.pvp || 0, 0,
              p.disponibleVenta ? "Sí" : "No", p.disponibleCompra ? "Sí" : "No"
            ]);
          } else {
            filas.push([p.nombre, p.categoria || "", v.sku, v.etiqueta || "Estándar", "Sin Bodega", 0]);
          }
          return;
        }

        bodegas.forEach(([nombreBodega, stockBodega]) => {
          if (esPro) {
            const limites = mapaLimites[v.sku + "|" + nombreBodega] || { minimo: 0, maximo: 0 };
            const sugerencia = (limites.maximo > stockBodega && stockBodega <= limites.minimo) ? (limites.maximo - stockBodega) : 0;
            filas.push([
              p.nombre, p.categoria || "", v.sku, v.etiqueta || "Estándar", nombreBodega, stockBodega,
              "Promedio Ponderado", v.costoPromedio || 0, limites.minimo, limites.maximo,
              calcularAlerta(stockBodega, limites), sugerencia, v.pvp || 0, (stockBodega * (v.costoPromedio || 0)),
              p.disponibleVenta ? "Sí" : "No", p.disponibleCompra ? "Sí" : "No"
            ]);
          } else {
            filas.push([p.nombre, p.categoria || "", v.sku, v.etiqueta || "Estándar", nombreBodega, stockBodega]);
          }
        });
      });
    });

    return filas;
  } catch (e) {
    console.error("Error al generar datos de exportación: " + e.message);
    return [["Error al generar el archivo"]];
  }
}

/**
 * Historial de movimientos de UNA variante (kardex detallado), incluyendo
 * los registros de Historial_Costos (compras de consumo interno que no
 * mueven stock pero si afectan costo).
 */
function obtenerHistorialMovimientosServidor(sku) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skuLower = sku.toString().trim().toLowerCase();
    const movimientos = [];
    const costoPromedioActual = (obtenerMapaCostosServidor()[sku] || 0);

    const hojaMov = ss.getSheetByName("Movimientos_Inventario");
    if (hojaMov) {
      const datos = hojaMov.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        if ((datos[i][2] || "").toString().trim().toLowerCase() === skuLower) {
          const motivo = datos[i][8] ? datos[i][8].toString() : "";
          const cantidad = parseFloat(datos[i][5]) || 0;
          const tipo = datos[i][3] ? datos[i][3].toString() : "";

          // Si el motivo trae "Costo unitario: $X" (entradas por compra reventa),
          // extraemos ese valor real. Para salidas o ajustes sin costo explicito,
          // usamos el costo promedio actual como referencia aproximada.
          let costoUnitario = costoPromedioActual;
          const match = motivo.match(/Costo unitario:\s*\$?([\d.]+)/);
          if (match) costoUnitario = parseFloat(match[1]) || costoPromedioActual;

          movimientos.push({
            idMovimiento: datos[i][0] ? datos[i][0].toString() : "",
            fecha: datos[i][1] ? datos[i][1].toString() : "",
            sku: datos[i][2] ? datos[i][2].toString() : "",
            tipo: tipo,
            origen: datos[i][4] ? datos[i][4].toString() : "",
            cantidad: cantidad,
            bodega: datos[i][6] ? datos[i][6].toString() : "",
            referencia: datos[i][7] ? datos[i][7].toString() : "",
            motivo: motivo,
            usuario: datos[i][9] ? datos[i][9].toString() : "",
            costoUnitario: costoUnitario,
            valorTotal: cantidad * costoUnitario
          });
        }
      }
    }

    const hojaHist = ss.getSheetByName("Historial_Costos");
    if (hojaHist) {
      const datosHist = hojaHist.getDataRange().getValues();
      for (let i = 1; i < datosHist.length; i++) {
        if ((datosHist[i][1] || "").toString().trim().toLowerCase() === skuLower) {
          const costoUnitario = parseFloat(datosHist[i][4]) || 0;
          const cantidad = parseFloat(datosHist[i][3]) || 0;
          movimientos.push({
            idMovimiento: "COSTO-" + i,
            fecha: datosHist[i][0] ? datosHist[i][0].toString() : "",
            sku: datosHist[i][1] ? datosHist[i][1].toString() : "",
            tipo: "Solo Costo",
            origen: datosHist[i][2] ? datosHist[i][2].toString() : "",
            cantidad: cantidad,
            bodega: "",
            referencia: "",
            motivo: datosHist[i][5] ? datosHist[i][5].toString() : "",
            usuario: "",
            costoUnitario: costoUnitario,
            valorTotal: cantidad * costoUnitario
          });
        }
      }
    }

    movimientos.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return movimientos;
  } catch (e) {
    console.error("Error al obtener historial de movimientos: " + e.message);
    return [];
  }
}

/**
 * Genera las filas (encabezado + datos) del historial de UNA variante,
 * listas para exportar a Excel desde el modal de Historial.
 */
function obtenerDatosExportacionHistorialServidor(sku) {
  try {
    const movs = obtenerHistorialMovimientosServidor(sku);
    const filas = [["Fecha", "SKU", "Tipo", "Origen", "Cantidad", "Bodega", "Costo Unitario", "Valor Total", "Motivo"]];

    // Para el Excel, en orden cronologico (mas antiguo primero) con total acumulado
    const cronologico = [...movs].reverse();
    let acumulado = 0;
    cronologico.forEach(m => {
      if (m.tipo !== "Solo Costo") {
        acumulado += (m.tipo === "Entrada") ? m.cantidad : -m.cantidad;
      }
      filas.push([
        m.fecha ? new Date(m.fecha).toLocaleString() : "",
        m.sku,
        m.tipo,
        m.origen,
        m.cantidad,
        m.bodega || "",
        m.costoUnitario || 0,
        m.valorTotal || 0,
        m.motivo || ""
      ]);
    });

    return filas;
  } catch (e) {
    console.error("Error al exportar historial: " + e.message);
    return [["Error al generar el archivo"]];
  }
}

/**
 * Genera el contenido completo para "Descargar Todo" desde el Historial:
 * una hoja "Global" (todos los movimientos, acumulado global) + una hoja
 * por cada bodega que tenga movimientos de esta variante (acumulado propio
 * de esa bodega). Los registros "Solo Costo" (compras de consumo interno)
 * solo aparecen en la hoja Global, porque no pertenecen a ninguna bodega.
 * @return {Object} { "Global": [[fila],[fila]...], "BodegaA": [...], ... }
 */
function obtenerDatosExportacionHistorialCompletoServidor(sku) {
  try {
    const movs = obtenerHistorialMovimientosServidor(sku);
    const encabezado = ["Fecha", "SKU", "Tipo", "Origen", "Cantidad", "Bodega", "Costo Unitario", "Valor Total", "Motivo"];

    function construirHoja(listaMovs) {
      const filas = [encabezado];
      const cronologico = [...listaMovs].reverse();
      let acumulado = 0;
      cronologico.forEach(m => {
        if (m.tipo !== "Solo Costo") {
          acumulado += (m.tipo === "Entrada") ? m.cantidad : -m.cantidad;
        }
        filas.push([
          m.fecha ? new Date(m.fecha).toLocaleString() : "",
          m.sku, m.tipo, m.origen, m.cantidad, m.bodega || "",
          m.costoUnitario || 0, m.valorTotal || 0, m.motivo || ""
        ]);
      });
      return filas;
    }

    const resultado = { "Global": construirHoja(movs) };

    const bodegasUsadas = [...new Set(movs.filter(m => m.bodega).map(m => m.bodega))];
    bodegasUsadas.forEach(bodega => {
      const movsDeEstaBodega = movs.filter(m => m.bodega === bodega);
      resultado[bodega] = construirHoja(movsDeEstaBodega);
    });

    return resultado;
  } catch (e) {
    console.error("Error al exportar historial completo: " + e.message);
    return { "Global": [["Error al generar el archivo"]] };
  }
}


// ==========================================
//   MODULO: USUARIOS, ROLES Y AUDITORIA
// ==========================================
//
// Identificacion: se usa el correo de Google de quien tiene la sesion abierta
// (Session.getActiveUser().getEmail()) en vez de un login propio con
// contraseña. El login propio queda pendiente para una fase futura.
//
// Hoja "Usuarios": Correo | Nombre | Departamento | Rol | Activo | Fecha Registro | Ultima Actividad
// Hoja "Roles": Nombre Rol | Permisos (JSON) | Fecha Creacion
//   Permisos JSON: { "Inventarios": {ver,crear,editar,eliminar}, "Compras": {...}, ... }
// Hoja "Departamentos": Nombre | Fecha Creacion
// Hoja "Auditoria": Fecha | Usuario | Modulo | Accion | Detalle | ValorAnterior | ValorNuevo

/**
 * Helper interno: registra un evento en la Bitacora de Auditoria. Es de solo
 * escritura (append), nunca se edita ni se borra una fila ya escrita —
 * eso es lo que hace que la bitacora sea confiable como registro historico.
 * Se llama desde las funciones de escritura de cualquier modulo.
 * @param {String} modulo Nombre del modulo (ej. "Inventarios", "Productos")
 * @param {String} accion "Crear" | "Editar" | "Eliminar"
 * @param {String} detalle Descripcion breve legible (ej. "Producto: Lapicero Azul")
 * @param {*} valorAnterior Opcional: el dato antes del cambio (se guarda como JSON)
 * @param {*} valorNuevo Opcional: el dato despues del cambio (se guarda como JSON)
 */
function registrarAuditoria(modulo, accion, detalle, valorAnterior, valorNuevo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Auditoria");
    if (!hoja) {
      hoja = ss.insertSheet("Auditoria");
      hoja.appendRow(["Fecha", "Usuario", "Modulo", "Accion", "Detalle", "Valor Anterior", "Valor Nuevo"]);
    }

    const correo = Session.getActiveUser().getEmail() || "Desconocido";

    hoja.appendRow([
      new Date(), correo, modulo, accion, detalle || "",
      valorAnterior !== undefined ? JSON.stringify(valorAnterior) : "",
      valorNuevo !== undefined ? JSON.stringify(valorNuevo) : ""
    ]);

    actualizarUltimaActividadServidor(correo);
  } catch (e) {
    console.error("Error al registrar auditoria: " + e.message);
  }
}

/**
 * Devuelve la bitacora completa, mas reciente primero. Acepta filtros
 * opcionales para la pantalla de Auditoria (por modulo o por usuario).
 * @param {Object} filtros Opcional: { modulo, usuario }
 */
function obtenerAuditoriaServidor(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Auditoria");
    if (!hoja) return [];

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const filtroModulo = (filtros && filtros.modulo) ? filtros.modulo.toString().trim() : "";
    const filtroUsuario = (filtros && filtros.usuario) ? filtros.usuario.toString().trim().toLowerCase() : "";

    const registros = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      const modulo = (datos[i][2] || "").toString();
      const usuario = (datos[i][1] || "").toString();

      if (filtroModulo && modulo !== filtroModulo) continue;
      if (filtroUsuario && usuario.toLowerCase() !== filtroUsuario) continue;

      registros.push({
        fecha: datos[i][0] ? datos[i][0].toString() : "",
        usuario: usuario,
        modulo: modulo,
        accion: datos[i][3] ? datos[i][3].toString() : "",
        detalle: datos[i][4] ? datos[i][4].toString() : "",
        valorAnterior: datos[i][5] ? datos[i][5].toString() : "",
        valorNuevo: datos[i][6] ? datos[i][6].toString() : ""
      });
    }

    registros.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return registros;
  } catch (e) {
    console.error("Error al obtener auditoria: " + e.message);
    return [];
  }
}

/**
 * Exporta la bitacora completa (o filtrada) a formato de filas para Excel.
 */
function obtenerDatosExportacionAuditoriaServidor(filtros) {
  try {
    const registros = obtenerAuditoriaServidor(filtros);
    const filas = [["Fecha", "Usuario", "Modulo", "Accion", "Detalle", "Valor Anterior", "Valor Nuevo"]];
    registros.forEach(r => {
      filas.push([
        r.fecha ? new Date(r.fecha).toLocaleString() : "",
        r.usuario, r.modulo, r.accion, r.detalle, r.valorAnterior, r.valorNuevo
      ]);
    });
    return filas;
  } catch (e) {
    console.error("Error al exportar auditoria: " + e.message);
    return [["Error al generar el archivo"]];
  }
}

// ------------------------------------------
//   USUARIOS
// ------------------------------------------

/**
 * Helper interno: actualiza la marca de "ultima actividad" de un usuario.
 * Se llama automaticamente desde registrarAuditoria, no hace falta invocarla aparte.
 */
function actualizarUltimaActividadServidor(correo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    if (!hoja) return;

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().toLowerCase() === correo.toLowerCase()) {
        hoja.getRange(i + 1, 7).setValue(new Date());
        return;
      }
    }
  } catch (e) {
    console.error("Error al actualizar ultima actividad: " + e.message);
  }
}

/**
 * Registra o actualiza un usuario del sistema.
 * @param {Object} datos { correo, nombre, departamento, rol, activo, esEdicion, correoOriginal }
 */
function registrarUsuarioServidor(datos) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Usuarios");
    if (!hoja) {
      hoja = ss.insertSheet("Usuarios");
      hoja.appendRow(["Correo", "Nombre", "Departamento", "Rol", "Activo", "Fecha Registro", "Ultima Actividad"]);
    }

    const correo = (datos.correo || "").toString().trim().toLowerCase();
    if (!correo) return { exito: false, mensaje: "El correo es obligatorio." };
    if (!datos.nombre) return { exito: false, mensaje: "El nombre es obligatorio." };

    const db = hoja.getDataRange().getValues();

    if (datos.esEdicion) {
      const correoOriginal = (datos.correoOriginal || "").toString().trim().toLowerCase();
      for (let i = 1; i < db.length; i++) {
        if ((db[i][0] || "").toString().trim().toLowerCase() === correoOriginal) {
          hoja.getRange(i + 1, 1, 1, 5).setValues([[correo, datos.nombre, datos.departamento || "", datos.rol || "", datos.activo !== false]]);
          registrarAuditoria("Usuarios", "Editar", "Usuario: " + datos.nombre, null, datos);
          return { exito: true, mensaje: "¡Usuario actualizado con éxito!" };
        }
      }
      return { exito: false, mensaje: "No se encontró el usuario original." };
    } else {
      for (let i = 1; i < db.length; i++) {
        if ((db[i][0] || "").toString().trim().toLowerCase() === correo) {
          return { exito: false, mensaje: "⚠️ Ya existe un usuario registrado con ese correo." };
        }
      }
      hoja.appendRow([correo, datos.nombre, datos.departamento || "", datos.rol || "", true, new Date(), ""]);
      registrarAuditoria("Usuarios", "Crear", "Usuario: " + datos.nombre, null, datos);
      return { exito: true, mensaje: "¡Usuario registrado con éxito!" };
    }
  });
}

function obtenerUsuariosServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const usuarios = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      usuarios.push({
        correo: datos[i][0].toString(),
        nombre: datos[i][1] ? datos[i][1].toString() : "",
        departamento: datos[i][2] ? datos[i][2].toString() : "",
        rol: datos[i][3] ? datos[i][3].toString() : "",
        activo: datos[i][4] !== false && datos[i][4] !== "FALSE",
        fechaRegistro: datos[i][5] ? datos[i][5].toString() : "",
        ultimaActividad: datos[i][6] ? datos[i][6].toString() : ""
      });
    }
    return usuarios;
  } catch (e) {
    console.error("Error al obtener usuarios: " + e.message);
    return [];
  }
}

/**
 * Helper global: resuelve el NOMBRE de un usuario a partir de su correo de
 * Google, cruzando contra la hoja Usuarios. Si no esta registrado (o no
 * tiene nombre), devuelve el correo tal cual como respaldo -- nunca deja el
 * campo vacio. Se usa en cualquier lugar donde se mostraba el correo crudo
 * y se quiere mostrar el nombre real de la persona (ej. "Registrado por").
 */
function resolverNombreUsuarioPorCorreoServidor(correo) {
  try {
    const correoLimpio = (correo || "").toString().trim().toLowerCase();
    if (!correoLimpio) return "";
    const usuarios = obtenerUsuariosServidor();
    const usuario = usuarios.find(u => u.correo.toString().trim().toLowerCase() === correoLimpio);
    return (usuario && usuario.nombre) ? usuario.nombre : correo;
  } catch (e) {
    return correo || "";
  }
}

/**
 * Activa o desactiva un usuario (no se elimina el registro, solo se inactiva,
 * para conservar el historial de auditoria asociado a ese correo).
 */
function cambiarEstadoUsuarioServidor(correo, activo) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    if (!hoja) return { exito: false, mensaje: "No existe el directorio de usuarios." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === correo.toString().trim().toLowerCase()) {
        hoja.getRange(i + 1, 5).setValue(!!activo);
        registrarAuditoria("Usuarios", "Editar", (activo ? "Activado: " : "Desactivado: ") + correo);
        return { exito: true, mensaje: activo ? "¡Usuario activado!" : "¡Usuario desactivado!" };
      }
    }
    return { exito: false, mensaje: "No se encontró el usuario." };
  });
}

/**
 * Devuelve los datos del usuario actual (segun su correo de sesion de Google),
 * incluyendo su rol y permisos ya resueltos. Si el correo no esta registrado
 * en Usuarios, devuelve un perfil "Invitado" sin permisos especiales.
 * Se llama al cargar cualquier modulo para saber que puede hacer esta persona.
 */
function obtenerPerfilActualServidor() {
  try {
    const correo = (Session.getActiveUser().getEmail() || "").toLowerCase();
    const usuarios = obtenerUsuariosServidor();
    const usuario = usuarios.find(u => u.correo.toLowerCase() === correo);

    if (!usuario || !usuario.activo) {
      return { correo: correo, nombre: "Invitado", departamento: "", rol: "", activo: false, permisos: {} };
    }

    const roles = obtenerRolesServidor();
    const rolInfo = roles.find(r => r.nombre === usuario.rol);

    return {
      correo: usuario.correo,
      nombre: usuario.nombre,
      departamento: usuario.departamento,
      rol: usuario.rol,
      activo: true,
      permisos: rolInfo ? rolInfo.permisos : {}
    };
  } catch (e) {
    console.error("Error al obtener perfil actual: " + e.message);
    return { correo: "", nombre: "Invitado", departamento: "", rol: "", activo: false, permisos: {} };
  }
}

// ------------------------------------------
//   ROLES Y PERMISOS (granularidad por modulo)
// ------------------------------------------

/**
 * Crea o actualiza un rol con sus permisos por modulo.
 * @param {Object} datos { nombre, permisos: { "Inventarios": {ver,crear,editar,eliminar}, ... }, esEdicion, nombreOriginal }
 */
function guardarRolServidor(datos) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Roles");
    if (!hoja) {
      hoja = ss.insertSheet("Roles");
      hoja.appendRow(["Nombre Rol", "Permisos (JSON)", "Fecha Creacion"]);
    }

    const nombre = (datos.nombre || "").toString().trim();
    if (!nombre) return { exito: false, mensaje: "El nombre del rol no puede estar vacío." };

    const db = hoja.getDataRange().getValues();

    if (datos.esEdicion) {
      const nombreOriginal = (datos.nombreOriginal || "").toString().trim();
      for (let i = 1; i < db.length; i++) {
        if ((db[i][0] || "").toString().trim() === nombreOriginal) {
          hoja.getRange(i + 1, 1, 1, 2).setValues([[nombre, JSON.stringify(datos.permisos || {})]]);
          registrarAuditoria("Roles", "Editar", "Rol: " + nombre);
          return { exito: true, mensaje: "¡Rol actualizado con éxito!" };
        }
      }
      return { exito: false, mensaje: "No se encontró el rol original." };
    } else {
      for (let i = 1; i < db.length; i++) {
        if ((db[i][0] || "").toString().trim().toLowerCase() === nombre.toLowerCase()) {
          return { exito: false, mensaje: "⚠️ Ya existe un rol con ese nombre." };
        }
      }
      hoja.appendRow([nombre, JSON.stringify(datos.permisos || {}), new Date()]);
      registrarAuditoria("Roles", "Crear", "Rol: " + nombre);
      return { exito: true, mensaje: "¡Rol creado con éxito!" };
    }
  });
}

function obtenerRolesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Roles");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const roles = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let permisos = {};
      try { permisos = JSON.parse(datos[i][1] || "{}"); } catch(e) {}
      roles.push({ nombre: datos[i][0].toString(), permisos: permisos });
    }
    return roles;
  } catch (e) {
    console.error("Error al obtener roles: " + e.message);
    return [];
  }
}

function eliminarRolServidor(nombre) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hojaUsuarios = ss.getSheetByName("Usuarios");
    if (hojaUsuarios) {
      const dbUsuarios = hojaUsuarios.getDataRange().getValues();
      for (let i = 1; i < dbUsuarios.length; i++) {
        if ((dbUsuarios[i][3] || "").toString() === nombre.toString()) {
          return { exito: false, mensaje: "⚠️ No se puede eliminar: hay usuarios con este rol asignado." };
        }
      }
    }

    const hoja = ss.getSheetByName("Roles");
    if (!hoja) return { exito: false, mensaje: "No existen roles registrados." };
    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim() === nombre.toString().trim()) {
        hoja.deleteRow(i + 1);
        registrarAuditoria("Roles", "Eliminar", "Rol: " + nombre);
        return { exito: true, mensaje: "¡Rol eliminado con éxito!" };
      }
    }
    return { exito: false, mensaje: "No se encontró el rol." };
  });
}

// ------------------------------------------
//   DEPARTAMENTOS
// ------------------------------------------

function crearDepartamentoServidor(nombre) {
  return ejecutarConLock(function() {
    const nombreLimpio = nombre.toString().trim();
    if (!nombreLimpio) return { exito: false, mensaje: "El nombre del departamento no puede estar vacío.", departamentos: obtenerDepartamentosServidor() };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Departamentos");
    if (!hoja) {
      hoja = ss.insertSheet("Departamentos");
      hoja.appendRow(["Nombre", "Fecha Creacion"]);
    }

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe un departamento con ese nombre.", departamentos: obtenerDepartamentosServidor() };
      }
    }

    hoja.appendRow(["'" + nombreLimpio, new Date()]);
    registrarAuditoria("Departamentos", "Crear", "Departamento: " + nombreLimpio);
    return { exito: true, mensaje: "¡Departamento creado con éxito!", departamentos: obtenerDepartamentosServidor() };
  });
}

function obtenerDepartamentosServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Departamentos");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    const departamentos = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") departamentos.push(datos[i][0].toString().trim());
    }
    return departamentos;
  } catch (e) {
    console.error("Error al obtener departamentos: " + e.message);
    return [];
  }
}

function eliminarDepartamentoServidor(nombre) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hojaUsuarios = ss.getSheetByName("Usuarios");
    if (hojaUsuarios) {
      const dbUsuarios = hojaUsuarios.getDataRange().getValues();
      for (let i = 1; i < dbUsuarios.length; i++) {
        if ((dbUsuarios[i][2] || "").toString() === nombre.toString()) {
          return { exito: false, mensaje: "⚠️ No se puede eliminar: hay usuarios en este departamento.", departamentos: obtenerDepartamentosServidor() };
        }
      }
    }

    const hoja = ss.getSheetByName("Departamentos");
    if (!hoja) return { exito: false, mensaje: "No existen departamentos registrados.", departamentos: [] };
    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombre.toString().trim().toLowerCase()) {
        hoja.deleteRow(i + 1);
        registrarAuditoria("Departamentos", "Eliminar", "Departamento: " + nombre);
        return { exito: true, mensaje: "¡Departamento eliminado con éxito!", departamentos: obtenerDepartamentosServidor() };
      }
    }
    return { exito: false, mensaje: "No se encontró el departamento.", departamentos: obtenerDepartamentosServidor() };
  });
}

// ==========================================
//        MODULO: PROVEEDORES
// ==========================================

function registrarProveedorServidor(datos) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Proveedores");

    if (!hoja) {
      hoja = ss.insertSheet("Proveedores");
      hoja.appendRow([
        "Nombre", "NIT/DUI", "NRC", "Contacto", "Telefono", "Correo", "Direccion",
        "Fecha Registro", "Tipo Contribuyente", "Tipo Persona", "WhatsApp", "Giro Comercial",
        "Categorias (JSON)", "Metodos Pago (JSON)", "Cuentas Bancarias (JSON)"
      ]);
    }

    const db = hoja.getDataRange().getValues();
    const tipoContribuyente = datos.tipoContribuyente || "Pequeño Contribuyente";
    const tipoPersona = datos.tipoPersona || "Persona Natural";
    const categorias = JSON.stringify(datos.categorias || []);
    const metodosPago = JSON.stringify(datos.metodosPago || []);
    const cuentasBancarias = JSON.stringify(datos.cuentasBancarias || []);

    // El NRC es obligatorio unicamente para proveedores tipo Empresa (Juridica).
    // Una Persona Natural puede tenerlo o no (comerciante individual inscrito).
    if (tipoPersona === "Empresa" && !(datos.nrc || "").toString().trim()) {
      return { exito: false, mensaje: "El NRC es obligatorio para proveedores tipo Empresa (Jurídica)." };
    }

    if (datos.esEdicion) {
      let filaDestino = -1;
      for (let i = 1; i < db.length; i++) {
        if (db[i][1].toString() === datos.nitDuiOriginal.toString()) {
          filaDestino = i + 1;
          break;
        }
      }
      if (filaDestino === -1) {
        return { exito: false, mensaje: "No se encontró el proveedor original para modificar." };
      }
      hoja.getRange(filaDestino, 1, 1, 7).setValues([[
        datos.nombre, "'" + datos.nitDui, "'" + (datos.nrc || ""), datos.contacto || "",
        datos.telefono || "", datos.correo || "", datos.direccion || ""
      ]]);
      hoja.getRange(filaDestino, 9, 1, 7).setValues([[tipoContribuyente, tipoPersona, datos.whatsapp || "", datos.giroComercial || "", categorias, metodosPago, cuentasBancarias]]);
      return { exito: true, mensaje: "¡Proveedor actualizado con éxito!" };
    } else {
      for (let i = 1; i < db.length; i++) {
        if (db[i][1].toString() === datos.nitDui.toString()) {
          return { exito: false, mensaje: "⚠️ Ya existe un proveedor registrado con ese NIT/DUI." };
        }
      }
      hoja.appendRow([
        datos.nombre, "'" + datos.nitDui, "'" + (datos.nrc || ""), datos.contacto || "",
        datos.telefono || "", datos.correo || "", datos.direccion || "", new Date(),
        tipoContribuyente, tipoPersona, datos.whatsapp || "", datos.giroComercial || "",
        categorias, metodosPago, cuentasBancarias
      ]);
      return { exito: true, mensaje: "¡Proveedor guardado con éxito!" };
    }
  });
}

function obtenerProveedoresServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Proveedores");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const proveedores = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let categorias = [];
      try { categorias = JSON.parse(datos[i][12] || "[]"); } catch(e) {}
      let metodosPago = [];
      try { metodosPago = JSON.parse(datos[i][13] || "[]"); } catch(e) {}
      let cuentasBancarias = [];
      try { cuentasBancarias = JSON.parse(datos[i][14] || "[]"); } catch(e) {}
      proveedores.push({
        nombre: datos[i][0].toString(),
        nitDui: datos[i][1] ? datos[i][1].toString() : "",
        nrc: datos[i][2] ? datos[i][2].toString() : "",
        contacto: datos[i][3] ? datos[i][3].toString() : "",
        telefono: datos[i][4] ? datos[i][4].toString() : "",
        correo: datos[i][5] ? datos[i][5].toString() : "",
        direccion: datos[i][6] ? datos[i][6].toString() : "",
        tipoContribuyente: datos[i][8] ? datos[i][8].toString() : "Pequeño Contribuyente",
        tipoPersona: datos[i][9] ? datos[i][9].toString() : "Persona Natural",
        whatsapp: datos[i][10] ? datos[i][10].toString() : "",
        giroComercial: datos[i][11] ? datos[i][11].toString() : "",
        categorias: categorias,
        metodosPago: metodosPago,
        cuentasBancarias: cuentasBancarias
      });
    }
    return proveedores;
  } catch (e) {
    console.error("Error al obtener proveedores: " + e.message);
    return [];
  }
}

function eliminarProveedorServidor(nitDui) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Proveedores");
    if (!hoja) return { exito: false, mensaje: "No existe el directorio de proveedores." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][1].toString() === nitDui.toString()) {
        hoja.deleteRow(i + 1);
        return { exito: true, mensaje: "¡Proveedor eliminado con éxito!" };
      }
    }
    return { exito: false, mensaje: "No se encontró el proveedor para eliminar." };
  });
}

// ------------------------------------------
//   CATEGORIAS DE PROVEEDOR (libres, creadas por el usuario)
// ------------------------------------------

function crearCategoriaProveedorServidor(nombre) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Categorias_Proveedor");
    if (!hoja) {
      hoja = ss.insertSheet("Categorias_Proveedor");
      hoja.appendRow(["Nombre Categoría", "Fecha Creación"]);
    }

    const nombreLimpio = (nombre || "").toString().trim();
    if (!nombreLimpio) {
      return { exito: false, mensaje: "El nombre de la categoría no puede estar vacío.", categorias: obtenerCategoriasProveedorServidor() };
    }

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe una categoría con ese nombre.", categorias: obtenerCategoriasProveedorServidor() };
      }
    }

    hoja.appendRow(["'" + nombreLimpio, new Date()]);
    return { exito: true, mensaje: "¡Categoría creada con éxito!", categorias: obtenerCategoriasProveedorServidor() };
  });
}

function obtenerCategoriasProveedorServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Categorias_Proveedor");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    const categorias = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") categorias.push(datos[i][0].toString().trim());
    }
    return categorias;
  } catch (e) {
    console.error("Error al obtener categorías de proveedor: " + e.message);
    return [];
  }
}

/**
 * Helper interno: revisa si algun proveedor tiene asignada esta categoria.
 */
function categoriaProveedorEstaEnUso(nombreCategoria) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Proveedores");
    if (!hoja) return false;
    const datos = hoja.getDataRange().getValues();
    const nombreLower = nombreCategoria.toString().trim().toLowerCase();

    for (let i = 1; i < datos.length; i++) {
      let categorias = [];
      try { categorias = JSON.parse(datos[i][12] || "[]"); } catch(e) {}
      if (categorias.map(c => c.toString().trim().toLowerCase()).includes(nombreLower)) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function eliminarCategoriaProveedorServidor(nombre) {
  return ejecutarConLock(function() {
    const nombreLimpio = (nombre || "").toString().trim();

    if (categoriaProveedorEstaEnUso(nombreLimpio)) {
      return { exito: false, mensaje: "⚠️ No se puede eliminar: hay proveedores que usan esta categoría.", categorias: obtenerCategoriasProveedorServidor() };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Categorias_Proveedor");
    if (!hoja) return { exito: false, mensaje: "No existen categorías registradas.", categorias: [] };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
        hoja.deleteRow(i + 1);
        return { exito: true, mensaje: "¡Categoría eliminada con éxito!", categorias: obtenerCategoriasProveedorServidor() };
      }
    }
    return { exito: false, mensaje: "No se encontró la categoría para eliminar.", categorias: obtenerCategoriasProveedorServidor() };
  });
}

// ==========================================
//        MODULO: COMPRAS
// ==========================================
//
// FLUJO: Requisicion (interna, sin proveedor) -> Orden de Compra (con
// proveedor y costos) -> Recepcion (parcial o total, mueve inventario real).
//
// Hoja "Requisiciones": ID | Fecha | Solicitante | Estado | Items(JSON) | Notas
//   Items: [{ sku, descripcion, cantidad }]
//   Estado: "Pendiente" | "Convertida a Orden" | "Cancelada"
//
// Hoja "Ordenes_Compra": ID | Fecha | ID_Requisicion | Proveedor | Estado | Items(JSON) | Notas
//   Items: [{ sku, descripcion, cantidadOrdenada, cantidadRecibida, costoUnitario }]
//   Estado: "Pendiente de Recepcion" | "Recibida Parcial" | "Recibida Completa" | "Cancelada"
//
// Hoja "Recepciones": ID | Fecha | ID_Orden | Items(JSON) | Usuario
//   Items: [{ sku, cantidadRecibida }]

/**
 * Crea una requisicion con uno o varios items. Si viene de un solo item
 * (ej. boton "Requisicion" en Inventarios), igual se envuelve en arreglo.
 * @param {Object} datos { solicitante, items: [{sku, descripcion, cantidad}], notas }
 */
function crearRequisicionServidor(datos) {
  return ejecutarConLock(function() {
    const items = datos.items || [];
    if (items.length === 0) {
      return { exito: false, mensaje: "La requisición necesita al menos un ítem." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Requisiciones");
    if (!hoja) {
      hoja = ss.insertSheet("Requisiciones");
      hoja.appendRow(["ID", "Fecha", "Solicitante", "Departamento", "Estado", "Items (JSON)", "Notas"]);
    }

    // Cada item arranca con su propio estado de seguimiento (uso PRO), un
    // comentario opcional individual, y cantidadComprada en 0 -- este ultimo
    // campo se va incrementando cada vez que una Orden de Compra toma una
    // porcion de este item, permitiendo compras parciales a proveedores
    // distintos hasta cubrir la cantidad total solicitada.
    const itemsConEstado = items.map(it => Object.assign({}, it, {
      estadoItem: "Pendiente de Compra",
      idOrdenAsociada: "",
      comentario: it.comentario || "",
      cantidadComprada: 0
    }));

    const idRequisicion = generarCorrelativoRequisicionServidor(hoja);
    const correo = Session.getActiveUser().getEmail() || "Usuario";

    hoja.appendRow([
      idRequisicion, new Date(),
      datos.solicitante || correo,
      datos.departamento || "",
      "Pendiente", JSON.stringify(itemsConEstado), datos.notas || ""
    ]);

    registrarAuditoria("Requisiciones", "Crear", "Requisición: " + idRequisicion);
    return { exito: true, mensaje: "¡Requisición creada con éxito!", idRequisicion: idRequisicion };
  });
}

/**
 * Genera un ID corto y legible tipo REQ-2026-001, correlativo por año.
 * Cuenta cuantas requisiciones ya existen con el mismo año en su ID.
 */
function generarCorrelativoRequisicionServidor(hoja) {
  try {
    const anioActual = new Date().getFullYear().toString();
    const db = hoja.getDataRange().getValues();
    let maxNumero = 0;

    for (let i = 1; i < db.length; i++) {
      const id = (db[i][0] || "").toString();
      const match = id.match(/^REQ-(\d{4})-(\d+)$/);
      if (match && match[1] === anioActual) {
        const numero = parseInt(match[2], 10);
        if (numero > maxNumero) maxNumero = numero;
      }
    }

    const siguienteNumero = (maxNumero + 1).toString().padStart(3, "0");
    return "REQ-" + anioActual + "-" + siguienteNumero;
  } catch (e) {
    console.error("Error al generar correlativo: " + e.message);
    return "REQ-" + new Date().getTime();
  }
}

/**
 * Devuelve las requisiciones visibles para QUIEN esta consultando:
 * - Si su rol tiene permiso "ver" en el modulo "Compras" (o es Administrador),
 *   ve TODAS las requisiciones de todos los departamentos.
 * - En caso contrario, solo ve las de SU PROPIO departamento.
 * Si la persona no esta registrada en Usuarios, ve unicamente lo que ella
 * misma creo (fallback seguro, nunca se le oculta su propio trabajo).
 */

/**
 * FUNCION DE DIAGNOSTICO TEMPORAL — ejecutar manualmente desde el editor de
 * Apps Script (seleccionar esta funcion en el menu desplegable de arriba y
 * darle "Ejecutar"). Muestra en un cuadro de dialogo exactamente que correos
 * estan guardados en Usuarios y cual es el correo de la sesion activa, para
 * encontrar la diferencia exacta sin revisar logs.
 */
/**
 * DIAGNOSTICO TEMPORAL: vuelca a una hoja el estado crudo de Limites_Inventario
 * completa, mas el resultado calculado de obtenerInventarioAgrupadoServidor
 * para CADA variante, para revisar por que una alerta de stock no se activa.
 * Se puede eliminar una vez resuelto.
 */
function diagnosticoLimitesInventario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDiag = ss.getSheetByName("DIAGNOSTICO_LIMITES");
  if (hojaDiag) ss.deleteSheet(hojaDiag);
  hojaDiag = ss.insertSheet("DIAGNOSTICO_LIMITES");

  hojaDiag.appendRow(["=== HOJA Limites_Inventario (CRUDA) ==="]);
  const hojaLimites = ss.getSheetByName("Limites_Inventario");
  if (hojaLimites) {
    const datosLimites = hojaLimites.getDataRange().getValues();
    datosLimites.forEach(fila => hojaDiag.appendRow(fila));
  } else {
    hojaDiag.appendRow(["(la hoja Limites_Inventario no existe todavia)"]);
  }

  hojaDiag.appendRow(["---"]);
  hojaDiag.appendRow(["=== RESULTADO CALCULADO por obtenerInventarioAgrupadoServidor (Vista Global) ==="]);
  hojaDiag.appendRow(["Producto", "SKU", "Etiqueta", "Stock", "Minimo", "Maximo", "Alerta", "SugerenciaCompra"]);

  const inventario = obtenerInventarioAgrupadoServidor(null);
  inventario.forEach(p => {
    (p.variantes || []).forEach(v => {
      hojaDiag.appendRow([p.nombre, v.sku, v.etiqueta || "", v.stock, v.minimo, v.maximo, v.alerta, v.sugerenciaCompra]);
    });
  });
}

/**
 * REPARACION TEMPORAL: limpia filas corruptas de Limites_Inventario que
 * quedaron con un formato viejo (desplazado), donde la columna "Bodega"
 * contiene en realidad un numero (el minimo de esa version vieja) en vez de
 * un nombre de bodega real o "Global". Esas filas duplicaban SKUs que ya
 * tienen su version correcta mas abajo en la hoja, y podian generar
 * ambiguedad en el mapa de limites. Se puede eliminar una vez ejecutada.
 */
function repararLimitesInventarioCorruptos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Limites_Inventario");
  if (!hoja) return "No existe la hoja Limites_Inventario.";

  const datos = hoja.getDataRange().getValues();
  const filasABorrar = [];

  for (let i = 1; i < datos.length; i++) {
    const valorBodega = (datos[i][1] || "").toString().trim();
    // Si la columna "Bodega" es puramente numerica, es una fila vieja
    // corrupta (formato anterior sin columna de bodega) -- una bodega real
    // siempre es texto (ej. "Global", "Bodega General").
    if (valorBodega !== "" && !isNaN(valorBodega)) {
      filasABorrar.push(i + 1);
    }
  }

  // Borrar de abajo hacia arriba para no desordenar los indices de fila
  filasABorrar.reverse().forEach(numFila => hoja.deleteRow(numFila));

  return "Filas corruptas eliminadas: " + filasABorrar.length;
}

function diagnosticoNombreSolicitante() {
  const correoSesion = Session.getActiveUser().getEmail();
  const usuarios = obtenerUsuariosServidor();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hojaDiag = ss.getSheetByName("DIAGNOSTICO_TEMP");
  if (hojaDiag) ss.deleteSheet(hojaDiag);
  hojaDiag = ss.insertSheet("DIAGNOSTICO_TEMP");

  let fila = 1;
  hojaDiag.getRange(fila++, 1).setValue("Correo de la sesion activa:");
  hojaDiag.getRange(fila++, 1).setValue("[" + correoSesion + "]");
  fila++;

  hojaDiag.getRange(fila++, 1).setValue("Usuarios registrados (" + usuarios.length + "):");
  usuarios.forEach(u => {
    hojaDiag.getRange(fila++, 1).setValue("Correo: [" + u.correo + "]");
    hojaDiag.getRange(fila - 1, 2).setValue("Nombre: " + u.nombre);
    hojaDiag.getRange(fila - 1, 3).setValue("Activo: " + u.activo);
  });
  fila++;

  const hoja = ss.getSheetByName("Requisiciones");
  if (hoja) {
    const datos = hoja.getDataRange().getValues();
    hojaDiag.getRange(fila++, 1).setValue("Ultimas requisiciones (columna Solicitante tal cual esta guardada):");
    for (let i = Math.max(1, datos.length - 3); i < datos.length; i++) {
      hojaDiag.getRange(fila++, 1).setValue("ID: " + datos[i][0]);
      hojaDiag.getRange(fila - 1, 2).setValue("Solicitante guardado: [" + datos[i][2] + "]");
    }
  }

  hojaDiag.autoResizeColumns(1, 3);
  SpreadsheetApp.flush();
}

function obtenerRequisicionesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisiciones");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const perfil = obtenerPerfilActualServidor();
    const vetTodo = perfil.rol === "Administrador" || (perfil.permisos && perfil.permisos["Compras"] && perfil.permisos["Compras"].ver);
    const usuarios = obtenerUsuariosServidor();

    function resolverNombreSolicitante(correo) {
      const correoLimpio = (correo || "").toString().trim().toLowerCase();
      const usuario = usuarios.find(u => u.correo.toString().trim().toLowerCase() === correoLimpio);
      return usuario ? usuario.nombre : correo;
    }

    let requisiciones = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][5] || "[]"); } catch(e) {}
      const correoSolicitante = datos[i][2] ? datos[i][2].toString() : "";
      requisiciones.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        solicitante: correoSolicitante,
        solicitanteNombre: resolverNombreSolicitante(correoSolicitante),
        departamento: datos[i][3] ? datos[i][3].toString() : "",
        estado: datos[i][4] ? datos[i][4].toString() : "",
        items: items,
        notas: datos[i][6] ? datos[i][6].toString() : ""
      });
    }

    if (!vetTodo) {
      if (perfil.activo && perfil.departamento) {
        requisiciones = requisiciones.filter(r => r.departamento === perfil.departamento);
      } else {
        requisiciones = requisiciones.filter(r => r.solicitante === perfil.correo);
      }
    }

    requisiciones.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return requisiciones;
  } catch (e) {
    console.error("Error al obtener requisiciones: " + e.message);
    return [];
  }
}

/**
 * Cancela uno o varios items especificos de una requisicion (no necesariamente
 * toda la requisicion). Cada item cancelado guarda su motivo. Si TODOS los
 * items terminan cancelados, la requisicion completa pasa a estado "Cancelada".
 * @param {Object} datos { idRequisicion, skus: [sku1, sku2...], motivo }
 */
function cancelarRequisicionServidor(datos) {
  return ejecutarConLock(function() {
    const idRequisicion = (datos.idRequisicion || datos).toString();
    const skusACancelar = datos.skus || [];
    const motivo = (datos.motivo || "Sin motivo especificado").toString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisiciones");
    if (!hoja) return { exito: false, mensaje: "No existen requisiciones registradas." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idRequisicion) continue;

      let items = [];
      try { items = JSON.parse(db[i][5] || "[]"); } catch(e) {}

      // Si no se especifican skus puntuales, se cancelan TODOS los items
      // pendientes (compatibilidad con cancelacion completa de toda la requisicion).
      const skusObjetivo = skusACancelar.length > 0 ? skusACancelar : items.map(it => it.sku);

      let algunoYaComprado = false;
      items.forEach(it => {
        if (!skusObjetivo.includes(it.sku)) return;
        if (it.estadoItem === "Comprado" || it.estadoItem === "Recibido") {
          algunoYaComprado = true;
          return;
        }
        it.estadoItem = "Cancelado";
        it.motivoCancelacion = motivo;
      });

      if (algunoYaComprado) {
        return { exito: false, mensaje: "⚠️ Algunos ítems seleccionados ya fueron comprados o recibidos y no se pueden cancelar." };
      }

      const todosCancelados = items.every(it => it.estadoItem === "Cancelado");
      const nuevoEstadoGeneral = todosCancelados ? "Cancelada" : (db[i][4] || "Pendiente").toString();

      hoja.getRange(i + 1, 5).setValue(nuevoEstadoGeneral);
      hoja.getRange(i + 1, 6).setValue(JSON.stringify(items));

      registrarAuditoria("Requisiciones", "Editar", "Ítems cancelados en " + idRequisicion + ": " + skusObjetivo.join(", ") + " | Motivo: " + motivo);
      return { exito: true, mensaje: todosCancelados ? "¡Requisición cancelada por completo!" : "¡Ítem(s) cancelado(s) con éxito!" };
    }
    return { exito: false, mensaje: "No se encontró la requisición." };
  });
}

/**
 * Elimina una requisicion (fila completa) o solo sus items en estado
 * "Cancelado" (modo parcial). Reglas de permiso:
 * - Si el rol del usuario tiene permiso "eliminar" en el modulo "Requisiciones",
 *   puede borrar sin importar el estado (asi se respeta la Matriz de Roles
 *   en vez de hardcodear un nombre de rol especifico).
 * - En caso contrario, solo puede borrar lo que YA esta cancelado: la
 *   requisicion completa si su estado general es "Cancelada", o items
 *   puntuales que esten en estado "Cancelado" dentro de una requisicion activa.
 * @param {Object} datos { idRequisicion, modo: "completa"|"items_cancelados" }
 */
function eliminarRequisicionServidor(datos) {
  return ejecutarConLock(function() {
    const idRequisicion = (datos.idRequisicion || "").toString();
    const modo = datos.modo || "completa";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisiciones");
    if (!hoja) return { exito: false, mensaje: "No existen requisiciones registradas." };

    const perfil = obtenerPerfilActualServidor();
    const puedeEliminarSinRestriccion = perfil.permisos && perfil.permisos["Requisiciones"] && perfil.permisos["Requisiciones"].eliminar;

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idRequisicion) continue;

      const estadoGeneral = (db[i][4] || "").toString();
      let items = [];
      try { items = JSON.parse(db[i][5] || "[]"); } catch(e) {}

      if (modo === "completa") {
        if (!puedeEliminarSinRestriccion && estadoGeneral !== "Cancelada") {
          return { exito: false, mensaje: "⚠️ Solo puedes eliminar requisiciones completamente canceladas." };
        }
        hoja.deleteRow(i + 1);
        registrarAuditoria("Requisiciones", "Eliminar", "Requisición eliminada por completo: " + idRequisicion);
        return { exito: true, mensaje: "¡Requisición eliminada con éxito!" };
      } else {
        // Modo "items_cancelados": quita del JSON solo los items en estado Cancelado.
        const itemsAEliminar = items.filter(it => it.estadoItem === "Cancelado");
        if (itemsAEliminar.length === 0) {
          return { exito: false, mensaje: "No hay ítems cancelados para eliminar en esta requisición." };
        }

        const nuevosItems = items.filter(it => it.estadoItem !== "Cancelado");

        if (nuevosItems.length === 0) {
          // Si no queda ningun item, se elimina la fila completa.
          hoja.deleteRow(i + 1);
          registrarAuditoria("Requisiciones", "Eliminar", "Requisición eliminada (sin ítems restantes): " + idRequisicion);
          return { exito: true, mensaje: "¡Requisición eliminada con éxito (no quedaban ítems activos)!" };
        }

        hoja.getRange(i + 1, 6).setValue(JSON.stringify(nuevosItems));
        registrarAuditoria("Requisiciones", "Eliminar", "Ítems cancelados eliminados de " + idRequisicion + ": " + itemsAEliminar.map(it => it.sku).join(", "));
        return { exito: true, mensaje: "¡Ítems cancelados eliminados con éxito!" };
      }
    }
    return { exito: false, mensaje: "No se encontró la requisición." };
  });
}

/**
 * Devuelve las requisiciones que TODAVIA tienen al menos un item con
 * cantidad pendiente de comprar (cantidad - cantidadComprada > 0, y no
 * cancelado). Se usa exclusivamente para el popup "Comprar desde
 * Requisición" en el modulo de Compras -- cada item trae su cantidadPendiente
 * ya calculada para que el comprador vea exactamente cuanto falta, incluso
 * si esa linea ya se compro parcialmente a otro proveedor antes.
 */
function obtenerRequisicionesPendientesDeComprarServidor() {
  try {
    const todas = obtenerRequisicionesServidor();
    const resultado = [];

    // Mapa sku -> tipoFiscal (Bienes/Servicios), resuelto desde el catalogo
    // de productos. Se necesita para que el selector de cargos/impuestos en
    // Compras sepa si un item de requisicion es un Bien o un Servicio (los
    // items de requisicion solo traen sku/descripcion/cantidad, no el tipo).
    const mapaTipoFiscal = {};
    obtenerProductosServidor().forEach(p => {
      const tipoFiscal = (p.tipo === 'Servicio') ? 'Servicios' : 'Bienes';
      (p.variantes || []).forEach(v => { mapaTipoFiscal[v.sku] = tipoFiscal; });
    });

    todas.forEach(r => {
      if (r.estado === 'Cancelada') return;

      const itemsPendientes = (r.items || [])
        .filter(it => it.estadoItem !== 'Cancelado')
        .map(it => {
          const comprada = parseFloat(it.cantidadComprada) || 0;
          const pendiente = (parseFloat(it.cantidad) || 0) - comprada;
          return Object.assign({}, it, { cantidadPendiente: pendiente, tipoFiscal: mapaTipoFiscal[it.sku] || 'Bienes' });
        })
        .filter(it => it.cantidadPendiente > 0);

      if (itemsPendientes.length > 0) {
        resultado.push({
          id: r.id,
          fecha: r.fecha,
          solicitanteNombre: r.solicitanteNombre,
          departamento: r.departamento,
          items: itemsPendientes
        });
      }
    });

    return resultado;
  } catch (e) {
    console.error("Error al obtener requisiciones pendientes de comprar: " + e.message);
    return [];
  }
}

/**
 * Crea una Orden de Compra. Puede venir de una Requisicion (idRequisicion)
 * o crearse directa (idRequisicion vacio). Cada item necesita su costoUnitario.
 * @param {Object} datos { idRequisicion, proveedor, items: [{sku,descripcion,cantidadOrdenada,costoUnitario}], notas }
 */
/**
 * Crea una Orden de Compra. Los items pueden venir de UNA O VARIAS
 * requisiciones distintas (compra masiva): cada item trae su propio
 * idRequisicionOrigen para poder actualizar el estado correcto en cada una.
 * Items sin idRequisicionOrigen son de una orden directa (sin requisicion previa).
 * @param {Object} datos { proveedor, items: [{sku,descripcion,cantidadOrdenada,costoUnitario,idRequisicionOrigen}], notas }
 */
/**
 * Motor de calculo fiscal para El Salvador, segun el Tipo de Documento que
 * el proveedor emitira y el Regimen configurado en Config_Empresa.
 *
 * Reglas aplicadas (simplificadas para MIPYMES, revisar con contador si el
 * caso es mas complejo):
 * - CCF (Comprobante de Credito Fiscal): el subtotal YA incluye IVA 13%
 *   desglosado. Si la empresa compradora es "GRANDE", se retiene 1% de IVA
 *   al proveedor (Retencion de IVA, Art. 162 inciso 2 CT).
 * - Factura (consumidor final): no se desglosa IVA por separado (va incluido
 *   en el precio). Si la empresa compradora es "GRANDE" y el proveedor NO es
 *   Gran Contribuyente, aplica Percepcion de IVA 1% adicional a pagar.
 * - Sujeto Excluido: no aplica IVA. Aplica Retencion de Renta 10% sobre el
 *   monto bruto (Art. 156 CT), no retencion de IVA.
 * @param {Array} items [{cantidadOrdenada, costoUnitario}, ...]
 * @param {String} tipoDocumento "CCF" | "Factura" | "Sujeto Excluido"
 * @param {String} nombreProveedor Para consultar su Tipo de Contribuyente
 */
function calcularFiscalOrdenCompraServidor(items, tipoDocumento, nombreProveedor) {
  try {
    const subtotal = items.reduce((acc, it) => acc + (parseFloat(it.cantidadOrdenada) || 0) * (parseFloat(it.costoUnitario) || 0), 0);

    const configEmpresa = obtenerConfiguracionEmpresaServidor();
    const regimenEmpresa = (configEmpresa && configEmpresa.regimen) ? configEmpresa.regimen : "INFORMAL";
    const esGranContribuyenteCompradora = regimenEmpresa === "GRANDE";

    const proveedores = obtenerProveedoresServidor();
    const proveedorInfo = proveedores.find(p => p.nombre === nombreProveedor);
    const proveedorEsGranContribuyente = proveedorInfo && proveedorInfo.tipoContribuyente === "Gran Contribuyente";

    let iva = 0, percepcionIva = 0, retencionIva = 0, retencionRenta = 0;

    if (tipoDocumento === "CCF") {
      iva = subtotal * 0.13;
      if (esGranContribuyenteCompradora) {
        retencionIva = subtotal * 0.01;
      }
    } else if (tipoDocumento === "Factura") {
      // El IVA va incluido en el precio (no se desglosa), pero si aplica
      // percepcion, esa SI se calcula sobre el monto facturado.
      if (esGranContribuyenteCompradora && !proveedorEsGranContribuyente) {
        percepcionIva = subtotal * 0.01;
      }
    } else if (tipoDocumento === "Sujeto Excluido") {
      retencionRenta = subtotal * 0.10;
    }

    const totalARetener = retencionIva + retencionRenta;
    const totalAPagar = subtotal + iva + percepcionIva - totalARetener;

    return {
      tipoDocumento: tipoDocumento,
      regimenEmpresaUsado: regimenEmpresa,
      proveedorEsGranContribuyente: !!proveedorEsGranContribuyente,
      subtotal: subtotal,
      iva: iva,
      percepcionIva: percepcionIva,
      retencionIva: retencionIva,
      retencionRenta: retencionRenta,
      totalARetener: totalARetener,
      totalAPagar: totalAPagar
    };
  } catch (e) {
    console.error("Error en calculo fiscal: " + e.message);
    return { tipoDocumento: tipoDocumento, subtotal: 0, iva: 0, percepcionIva: 0, retencionIva: 0, retencionRenta: 0, totalARetener: 0, totalAPagar: 0 };
  }
}

/**
 * Crea una Orden de Compra bajo el nuevo modelo:
 * - Cada item lleva sus propios cargos/impuestos (array de {nombre, porcentaje}),
 *   en vez de un tipo de documento unico para toda la orden.
 * - Cada item puede (opcionalmente) traer idRequisicionOrigen + skuOrigen para
 *   saber de que linea de que requisicion proviene. Al guardar la orden, se
 *   SUMA la cantidad comprada en esa linea especifica de la requisicion
 *   (cantidadComprada += cantidadOrdenada), permitiendo compras parciales de
 *   una misma linea a proveedores distintos en momentos distintos.
 * - estadoPago arranca siempre en "Pendiente" (independiente del estado de
 *   recepcion, que sigue siendo "Pendiente de Recepción").
 * @param {Object} datos {
 *   proveedor, origenTipo: "Requisición"|"Directa", bodegaDestino, notas,
 *   items: [{ sku, descripcion, cantidadOrdenada, costoUnitario, cargos: [{nombre,porcentaje}],
 *             idRequisicionOrigen, skuOrigen }]
 * }
 */
/**
 * Genera un ID corto y legible tipo OC-2026-001, correlativo por año, igual
 * al patron usado en Requisiciones.
 */
function generarCorrelativoOrdenCompraServidor(hoja) {
  try {
    const anioActual = new Date().getFullYear().toString();
    const db = hoja.getDataRange().getValues();
    let maxNumero = 0;

    for (let i = 1; i < db.length; i++) {
      const id = (db[i][0] || "").toString();
      const match = id.match(/^OC-(\d{4})-(\d+)$/);
      if (match && match[1] === anioActual) {
        const numero = parseInt(match[2], 10);
        if (numero > maxNumero) maxNumero = numero;
      }
    }

    const siguienteNumero = (maxNumero + 1).toString().padStart(3, "0");
    return "OC-" + anioActual + "-" + siguienteNumero;
  } catch (e) {
    console.error("Error al generar correlativo de orden: " + e.message);
    return "OC-" + new Date().getTime();
  }
}

/**
 * Genera un correlativo corto tipo WH-IN-001 (entradas/recepciones) o
 * WH-OUT-001 (salidas), contando solo dentro de su propio prefijo -- no
 * lleva año, simplemente sigue creciendo. Se usa en la hoja "Recepciones"
 * para identificar tanto entradas (mercaderia recibida) como salidas
 * (entregas a Produccion u otras areas), distinguibles a simple vista.
 */
function generarCorrelativoRecepcionServidor(hoja, prefijo) {
  try {
    const db = hoja.getDataRange().getValues();
    let maxNumero = 0;
    const regex = new RegExp("^" + prefijo + "-(\\d+)$");

    for (let i = 1; i < db.length; i++) {
      const id = (db[i][0] || "").toString();
      const match = id.match(regex);
      if (match) {
        const numero = parseInt(match[1], 10);
        if (numero > maxNumero) maxNumero = numero;
      }
    }

    const siguienteNumero = (maxNumero + 1).toString().padStart(3, "0");
    return prefijo + "-" + siguienteNumero;
  } catch (e) {
    console.error("Error al generar correlativo de recepcion: " + e.message);
    return prefijo + "-" + new Date().getTime();
  }
}

function crearOrdenCompraServidor(datos) {
  return ejecutarConLock(function() {
    const items = datos.items || [];
    if (items.length === 0) {
      return { exito: false, mensaje: "La orden necesita al menos un ítem." };
    }
    if (!datos.proveedor) {
      return { exito: false, mensaje: "Selecciona o escribe un proveedor." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) {
      hoja = ss.insertSheet("Ordenes_Compra");
      hoja.appendRow([
        "ID", "Fecha", "ID_Requisicion", "Proveedor", "Estado", "Items (JSON)", "Notas",
        "Origen Tipo", "Bodega Destino", "Estado Pago", "Hitos Pago (JSON)",
        "Tipo Destino", "Entrega Directa A", "Fecha Último Envío", "Términos de Pago", "Llegada Prevista"
      ]);
    }

    // Cada item se guarda con cantidadRecibida en 0 (recepcion) y los
    // cargos que el comprador le asigno (puede ser un arreglo vacio si no
    // lleva ningun cargo). El costo total por item se calcula en el cliente
    // para mostrarlo, pero aqui se guarda el detalle crudo para recalcular
    // siempre desde la fuente de verdad.
    const itemsGuardados = items.map(it => Object.assign({}, it, {
      cantidadRecibida: 0,
      cargos: it.cargos || []
    }));

    const idOrden = generarCorrelativoOrdenCompraServidor(hoja);
    const origenTipo = datos.origenTipo || "Directa";
    const tipoDestino = (datos.tipoDestino === "Entrega Directa") ? "Entrega Directa" : "Bodega";

    // ID_Requisicion guarda TODAS las requisiciones de origen involucradas
    // (puede ser mas de una si se combinan items de varias requisiciones en
    // una sola orden), separadas por coma. Una orden "Directa" no tiene ninguna.
    const idsRequisicionesOrigen = [...new Set(items.filter(it => it.idRequisicionOrigen).map(it => it.idRequisicionOrigen))];

    // Si la orden es totalmente Directa (ningun item viene de requisicion),
    // se anota una nota visible aclarando que fue una compra sin requisicion
    // previa, para que quede explicito en el registro.
    const notasFinal = idsRequisicionesOrigen.length === 0
      ? "[Compra sin requisición previa] " + (datos.notas || "")
      : (datos.notas || "");

    hoja.appendRow([
      idOrden, new Date(), idsRequisicionesOrigen.join(","), datos.proveedor,
      "Pendiente de Recepción", JSON.stringify(itemsGuardados), notasFinal,
      origenTipo, datos.bodegaDestino || "", "Pendiente", "[]",
      tipoDestino, datos.entregaDirectaA || "", "",
      datos.terminosPago || "Contado", datos.llegadaPrevista || ""
    ]);

    // Por cada item que vino de una requisicion, suma la cantidad comprada
    // en ESA linea especifica (identificada por skuOrigen, que puede diferir
    // del sku final si el comprador sustituyo el producto). Esto permite que
    // una misma linea de requisicion se compre en pedazos a proveedores
    // distintos sin perder el rastro de cuanto falta.
    if (idsRequisicionesOrigen.length > 0) {
      const hojaReq = ss.getSheetByName("Requisiciones");
      if (hojaReq) {
        const dbReq = hojaReq.getDataRange().getValues();
        for (let i = 1; i < dbReq.length; i++) {
          const idReq = dbReq[i][0].toString();
          if (!idsRequisicionesOrigen.includes(idReq)) continue;

          let itemsReq = [];
          try { itemsReq = JSON.parse(dbReq[i][5] || "[]"); } catch(e) {}

          const itemsDeEstaReqEnEstaOrden = items.filter(it => it.idRequisicionOrigen === idReq);

          itemsDeEstaReqEnEstaOrden.forEach(itemComprado => {
            const skuBuscado = itemComprado.skuOrigen || itemComprado.sku;
            const itemReq = itemsReq.find(it => it.sku === skuBuscado);
            if (!itemReq) return;

            const comprada = (parseFloat(itemReq.cantidadComprada) || 0) + (parseFloat(itemComprado.cantidadOrdenada) || 0);
            itemReq.cantidadComprada = comprada;

            const totalSolicitado = parseFloat(itemReq.cantidad) || 0;
            itemReq.estadoItem = comprada >= totalSolicitado ? "Comprado" : "Comprado Parcial";

            // Se ACUMULAN todas las ordenes de compra que han tomado parte de
            // esta linea (no se sobreescribe) -- asi, si una misma linea se
            // compra en pedazos a proveedores distintos en momentos distintos,
            // quien creo la requisicion (y no tiene acceso al modulo de
            // Compras) puede ver TODOS los numeros de orden para consultar,
            // no solo el de la ultima compra parcial.
            if (!Array.isArray(itemReq.ordenesAsociadas)) {
              itemReq.ordenesAsociadas = itemReq.idOrdenAsociada ? [itemReq.idOrdenAsociada] : [];
            }
            if (!itemReq.ordenesAsociadas.includes(idOrden)) itemReq.ordenesAsociadas.push(idOrden);
            itemReq.idOrdenAsociada = idOrden; // se conserva por compatibilidad, ahora representa solo la mas reciente
          });

          const todosComprados = itemsReq.every(it => it.estadoItem === "Comprado" || it.estadoItem === "Cancelado");
          hojaReq.getRange(i + 1, 6).setValue(JSON.stringify(itemsReq));
          if (todosComprados) {
            hojaReq.getRange(i + 1, 5).setValue("Convertida a Orden");
          }
        }
      }
    }

    registrarAuditoria("Compras", "Crear", "Orden de Compra: " + idOrden);
    return { exito: true, mensaje: "¡Orden de Compra " + idOrden + " creada con éxito!", idOrden: idOrden };
  });
}

/**
 * Envia un correo de solicitud de cotizacion a varios proveedores a la vez,
 * todos en copia oculta (CCO) entre si, para que ninguno vea a los demas.
 * Es solo texto (sin precios, sin PDF): el proveedor responde con su oferta.
 * @param {Object} datos { correosProveedores: [...], items: [{descripcion,cantidad}] }
 */
function enviarCotizacionProveedoresServidor(datos) {
  try {
    const correos = (datos.correosProveedores || []).filter(c => c && c.trim() !== "");
    const items = datos.items || [];

    if (correos.length === 0) return { exito: false, mensaje: "Selecciona al menos un proveedor con correo registrado." };
    if (items.length === 0) return { exito: false, mensaje: "No hay ítems para cotizar." };

    const configEmpresa = obtenerConfiguracionEmpresaServidor();
    const nombreEmpresa = (configEmpresa && configEmpresa.nombreEmpresa) ? configEmpresa.nombreEmpresa : "Nuestra empresa";

    const listaItems = items.map(it => "- " + it.descripcion + " (Cantidad: " + it.cantidad + ")").join("\n");

    const asunto = "Solicitud de Cotización - " + nombreEmpresa;
    const cuerpo =
      "Estimado proveedor,\n\n" +
      "Solicitamos su cotización para los siguientes artículos:\n\n" +
      listaItems +
      "\n\nAgradecemos nos indique precio unitario, tiempo de entrega y condiciones de pago.\n\n" +
      "Quedamos atentos a su respuesta.\n\n" +
      "Saludos,\n" + nombreEmpresa;

    // El primer correo va como destinatario "para" (requerido por MailApp),
    // y el resto en CCO junto con el mismo primero, asi nadie ve al resto.
    MailApp.sendEmail({
      to: correos[0],
      bcc: correos.join(","),
      subject: asunto,
      body: cuerpo
    });

    registrarAuditoria("Compras", "Crear", "Cotización enviada a " + correos.length + " proveedor(es).");
    return { exito: true, mensaje: "¡Cotización enviada a " + correos.length + " proveedor(es) con éxito!" };
  } catch (error) {
    return { exito: false, mensaje: "Error al enviar cotización: " + error.message };
  }
}

/**
 * Envia el detalle de una Orden de Compra ya creada al correo del proveedor.
 * Funcion PRO -- el usuario elige enviarla desde el Registro de Compras.
 * @param {Object} datos { idOrden, correoProveedor }
 */
function enviarOrdenCompraCorreoServidor(datos) {
  try {
    const correoProveedor = (datos.correoProveedor || "").toString().trim();
    if (!correoProveedor) return { exito: false, mensaje: "El proveedor no tiene correo registrado." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen órdenes registradas." };

    const ordenes = obtenerOrdenesCompraServidor();
    const orden = ordenes.find(o => o.id === datos.idOrden);
    if (!orden) return { exito: false, mensaje: "No se encontró la orden de compra." };

    const configEmpresa = obtenerConfiguracionEmpresaServidor();
    const nombreEmpresa = (configEmpresa && configEmpresa.nombreEmpresa) ? configEmpresa.nombreEmpresa : "Nuestra empresa";

    const listaItems = (orden.items || []).map(it => "- " + it.descripcion + " (Cantidad: " + it.cantidadOrdenada + ")").join("\n");
    const asunto = "Orden de Compra " + orden.id + " - " + nombreEmpresa;
    const cuerpo =
      "Estimado proveedor,\n\n" +
      "Le compartimos el detalle de la Orden de Compra " + orden.id + ":\n\n" +
      listaItems +
      "\n\nTotal a pagar: $" + (orden.totalAPagar || 0).toFixed(2) +
      "\n\nQuedamos atentos a su confirmación.\n\n" +
      "Saludos,\n" + nombreEmpresa;

    const opcionesCorreo = { to: correoProveedor, subject: asunto, body: cuerpo };

    // El PDF se genera en el navegador (jsPDF) y llega aqui como base64 para
    // adjuntarlo realmente al correo -- MailApp.sendEmail acepta adjuntos
    // via Utilities.newBlob a partir de datos base64.
    if (datos.pdfBase64) {
      try {
        const bytes = Utilities.base64Decode(datos.pdfBase64);
        const blob = Utilities.newBlob(bytes, "application/pdf", orden.id + ".pdf");
        opcionesCorreo.attachments = [blob];
      } catch (eAdjunto) {
        console.error("No se pudo adjuntar el PDF: " + eAdjunto.message);
      }
    }

    MailApp.sendEmail(opcionesCorreo);

    // Registra la fecha del ultimo envio en la orden. El mismo boton de
    // "Enviar" sirve para reenviar -- no hay un boton separado de reenvio,
    // simplemente se actualiza esta fecha cada vez que se usa.
    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() === orden.id) {
        hoja.getRange(i + 1, 14).setValue(new Date());
        break;
      }
    }

    registrarAuditoria("Compras", "Editar", "Orden de Compra " + orden.id + " enviada por correo a " + correoProveedor);
    return { exito: true, mensaje: "¡Orden enviada por correo con éxito!" };
  } catch (error) {
    return { exito: false, mensaje: "Error al enviar la orden: " + error.message };
  }
}

/**
 * Devuelve solo las Ordenes de Compra de UN proveedor especifico (por nombre
 * exacto). Se usa en la ficha de Proveedores (funcion PRO) para mostrar el
 * conteo de ordenes asignadas y poder revisarlas al hacer clic.
 */
// ------------------------------------------
//   CARGOS / IMPUESTOS DE COMPRA (catalogo configurable)
// ------------------------------------------
//
// Reemplaza el motor fiscal fijo (CCF/Factura/Sujeto Excluido) por un
// catalogo libre que el usuario gestiona: cada cargo tiene un nombre y un
// porcentaje, y se aplica como "badge" a nivel de CADA ITEM dentro de una
// orden de compra (un item puede llevar varios cargos a la vez, ej. IVA +
// Retencion Renta). Esto da flexibilidad total sin asumir reglas fiscales
// que no encajen con la realidad de cada proveedor/categoria.

function crearCargoCompraServidor(datos) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Cargos_Compra");
    if (!hoja) {
      hoja = ss.insertSheet("Cargos_Compra");
      hoja.appendRow(["Nombre", "Porcentaje", "Tipo", "Orden", "Fecha Creación", "Tipo Compra", "Tipo Proveedor"]);
    }

    const nombre = (datos.nombre || "").toString().trim();
    const porcentaje = parseFloat(datos.porcentaje);
    const tipo = (datos.tipo === "Resta") ? "Resta" : "Suma";
    const tipoCompraValidos = ["Bienes", "Servicios", "Ambos"];
    const tipoCompra = tipoCompraValidos.includes(datos.tipoCompra) ? datos.tipoCompra : "Ambos";
    const tipoProveedorValidos = ["Persona Natural", "Empresa", "Ambos"];
    const tipoProveedor = tipoProveedorValidos.includes(datos.tipoProveedor) ? datos.tipoProveedor : "Ambos";

    if (!nombre) return { exito: false, mensaje: "El nombre del cargo no puede estar vacío.", cargos: obtenerCargosCompraServidor() };
    if (isNaN(porcentaje) || porcentaje < 0) return { exito: false, mensaje: "El porcentaje debe ser un número válido.", cargos: obtenerCargosCompraServidor() };

    const db = hoja.getDataRange().getValues();
    let maxOrden = 0;
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombre.toLowerCase()) {
        return { exito: false, mensaje: "⚠️ Ya existe un cargo con ese nombre.", cargos: obtenerCargosCompraServidor() };
      }
      const ordenFila = parseInt(db[i][3], 10) || 0;
      if (ordenFila > maxOrden) maxOrden = ordenFila;
    }

    // El orden se asigna automaticamente: el siguiente disponible despues
    // del mayor existente, asi cada cargo nuevo se agrega al final de la
    // secuencia de visualizacion. Se puede reordenar despues con
    // reordenarCargoCompraServidor.
    hoja.appendRow([nombre, porcentaje, tipo, maxOrden + 1, new Date(), tipoCompra, tipoProveedor]);
    return { exito: true, mensaje: "¡Cargo creado con éxito!", cargos: obtenerCargosCompraServidor() };
  });
}

function obtenerCargosCompraServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cargos_Compra");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    const cargos = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim() !== "") {
        cargos.push({
          nombre: datos[i][0].toString().trim(),
          porcentaje: parseFloat(datos[i][1]) || 0,
          tipo: (datos[i][2] || "Suma").toString().trim() === "Resta" ? "Resta" : "Suma",
          orden: parseInt(datos[i][3], 10) || 0,
          tipoCompra: (datos[i][5] || "Ambos").toString().trim() || "Ambos",
          tipoProveedor: (datos[i][6] || "Ambos").toString().trim() || "Ambos"
        });
      }
    }
    cargos.sort((a, b) => a.orden - b.orden);
    return cargos;
  } catch (e) {
    console.error("Error al obtener cargos de compra: " + e.message);
    return [];
  }
}

function eliminarCargoCompraServidor(nombre) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cargos_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen cargos registrados.", cargos: [] };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if ((db[i][0] || "").toString().trim().toLowerCase() === nombre.toString().trim().toLowerCase()) {
        hoja.deleteRow(i + 1);
        return { exito: true, mensaje: "¡Cargo eliminado con éxito!", cargos: obtenerCargosCompraServidor() };
      }
    }
    return { exito: false, mensaje: "No se encontró el cargo para eliminar.", cargos: obtenerCargosCompraServidor() };
  });
}

/**
 * Mueve un cargo una posicion arriba o abajo en el orden de visualizacion,
 * intercambiando su numero de orden con el del cargo vecino. No afecta el
 * calculo (todos los cargos siguen calculandose sobre el subtotal original
 * de cada item, en paralelo) -- el orden es puramente visual/de presentacion.
 * @param {String} nombre
 * @param {String} direccion "arriba" | "abajo"
 */
function reordenarCargoCompraServidor(nombre, direccion) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cargos_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen cargos registrados.", cargos: [] };

    const db = hoja.getDataRange().getValues();
    const filas = [];
    for (let i = 1; i < db.length; i++) {
      if (db[i][0] && db[i][0].toString().trim() !== "") {
        filas.push({ fila: i + 1, nombre: db[i][0].toString().trim(), orden: parseInt(db[i][3], 10) || 0 });
      }
    }
    filas.sort((a, b) => a.orden - b.orden);

    const idx = filas.findIndex(f => f.nombre.toLowerCase() === nombre.toString().trim().toLowerCase());
    if (idx === -1) return { exito: false, mensaje: "No se encontró el cargo.", cargos: obtenerCargosCompraServidor() };

    const idxVecino = direccion === "arriba" ? idx - 1 : idx + 1;
    if (idxVecino < 0 || idxVecino >= filas.length) {
      return { exito: false, mensaje: "Ya está en el extremo de la lista.", cargos: obtenerCargosCompraServidor() };
    }

    const ordenActual = filas[idx].orden;
    const ordenVecino = filas[idxVecino].orden;
    hoja.getRange(filas[idx].fila, 4).setValue(ordenVecino);
    hoja.getRange(filas[idxVecino].fila, 4).setValue(ordenActual);

    return { exito: true, mensaje: "¡Orden actualizado!", cargos: obtenerCargosCompraServidor() };
  });
}

function obtenerOrdenesPorProveedorServidor(nombreProveedor) {
  try {
    const todas = obtenerOrdenesCompraServidor();
    const nombreLower = (nombreProveedor || "").toString().trim().toLowerCase();
    return todas.filter(o => (o.proveedor || "").toString().trim().toLowerCase() === nombreLower);
  } catch (e) {
    console.error("Error al obtener órdenes por proveedor: " + e.message);
    return [];
  }
}

function obtenerOrdenesCompraServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const ordenes = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][5] || "[]"); } catch(e) {}
      let hitosPago = [];
      try { hitosPago = JSON.parse(datos[i][10] || "[]"); } catch(e) {}

      const totalAPagar = calcularTotalAPagarOrden(items);
      const totalPagado = hitosPago.reduce((acc, h) => acc + (parseFloat(h.monto) || 0), 0);

      ordenes.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        idRequisicion: datos[i][2] ? datos[i][2].toString() : "",
        proveedor: datos[i][3] ? datos[i][3].toString() : "",
        estado: datos[i][4] ? datos[i][4].toString() : "",
        items: items,
        notas: datos[i][6] ? datos[i][6].toString() : "",
        origenTipo: datos[i][7] ? datos[i][7].toString() : "Directa",
        bodegaDestino: datos[i][8] ? datos[i][8].toString() : "",
        estadoPago: datos[i][9] ? datos[i][9].toString() : "Pendiente",
        totalAPagar: totalAPagar,
        hitosPago: hitosPago,
        totalPagado: totalPagado,
        saldoPendiente: Math.max(0, totalAPagar - totalPagado),
        tipoDestino: datos[i][11] ? datos[i][11].toString() : "Bodega",
        entregaDirectaA: datos[i][12] ? datos[i][12].toString() : "",
        fechaUltimoEnvio: datos[i][13] ? datos[i][13].toString() : "",
        terminosPago: datos[i][14] ? datos[i][14].toString() : "Contado",
        llegadaPrevista: formatearFechaSimple(datos[i][15])
      });
    }
    ordenes.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return ordenes;
  } catch (e) {
    console.error("Error al obtener órdenes de compra: " + e.message);
    return [];
  }
}

function cancelarOrdenCompraServidor(idOrden) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen órdenes registradas." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idOrden.toString()) continue;

      const estado = (db[i][4] || "").toString();
      if (estado === "Recibida Completa" || estado === "Recibida Parcial") {
        return { exito: false, mensaje: "⚠️ No se puede cancelar: ya tiene mercadería recibida." };
      }

      // Cancelar la orden
      hoja.getRange(i + 1, 5).setValue("Cancelada");

      // Reactivar los items de las requisiciones de origen: cualquier item
      // que estaba vinculado a esta OC (idRequisicionOrigen o idOrdenAsociada)
      // vuelve a "Pendiente" para que pueda asignarse a una nueva orden.
      const idsReqOrigen = (db[i][2] || "").toString().split(",").map(s => s.trim()).filter(Boolean);
      let itemsReactivados = 0;

      if (idsReqOrigen.length > 0) {
        const hojaReq = ss.getSheetByName("Requisiciones");
        if (hojaReq) {
          const dbReq = hojaReq.getDataRange().getValues();
          for (let j = 1; j < dbReq.length; j++) {
            const idReq = dbReq[j][0].toString();
            if (!idsReqOrigen.includes(idReq)) continue;

            let items = [];
            try { items = JSON.parse(dbReq[j][5] || "[]"); } catch(e) {}

            let cambio = false;
            items.forEach(it => {
              // Solo reactiva items que estaban vinculados a ESTA orden
              // especificamente -- si el item ya fue recibido (estadoItem
              // "Recibido") no se toca, porque la mercaderia ya llego.
              if (it.idOrdenAsociada === idOrden && it.estadoItem !== "Recibido") {
                it.estadoItem = "Pendiente de Compra";
                it.idOrdenAsociada = "";
                cambio = true;
                itemsReactivados++;
              }
            });

            if (cambio) {
              hojaReq.getRange(j + 1, 6).setValue(JSON.stringify(items));

              // Si la requisicion estaba "En Proceso" y ahora todos sus items
              // estan Pendientes, regresa a estado "Pendiente" tambien.
              const estadoReq = (dbReq[j][3] || "").toString();
              if (estadoReq === "En Proceso") {
                const todosLibres = items.every(it => it.estadoItem === "Pendiente de Compra" || !it.estadoItem);
                if (todosLibres) {
                  hojaReq.getRange(j + 1, 4).setValue("Pendiente");
                }
              }
            }
          }
        }
      }

      const msgExtra = itemsReactivados > 0
        ? ` ${itemsReactivados} ítem(s) de requisición reactivados y disponibles para una nueva orden.`
        : "";

      registrarAuditoria("Compras", "Cancelar", "Orden cancelada: " + idOrden);
      return { exito: true, mensaje: "¡Orden de Compra cancelada!" + msgExtra };
    }
    return { exito: false, mensaje: "No se encontró la orden." };
  });
}

/**
 * Actualiza el estado de PAGO de una orden (independiente del estado de
 * RECEPCION). Valores esperados: "Pendiente", "Programado", "Parcial", "Pagada".
 * Es solo una etiqueta de seguimiento para el comprador -- no mueve dinero
 * ni genera movimientos contables, simplemente ayuda a saber que falta
 * programar o pagar a cada proveedor.
 */
/**
 * Registra un HITO de pago (anticipo, contado, saldo, etc.) sobre una orden
 * de compra existente. Varios hitos pueden coexistir en la misma orden
 * (ej. Anticipo $50 + Saldo $61.20), manteniendo trazabilidad completa sin
 * necesidad de abrir una orden nueva por cada pago parcial.
 *
 * Valida que la suma de TODOS los hitos (los ya existentes + el nuevo) no
 * supere el total a pagar de la orden -- evita pagar de mas por error.
 *
 * El estado de pago de la orden ("Pendiente"/"Parcial"/"Pagada") se
 * recalcula automaticamente a partir de la suma de hitos, en vez de
 * elegirse manualmente: si la suma cubre el 100% del total, pasa a "Pagada";
 * si es mayor a 0 pero no llega al 100%, "Parcial"; si es 0, "Pendiente".
 *
 * @param {Object} datos { idOrden, monto, fecha, metodo, nota }
 */
function registrarHitoPagoOrdenServidor(datos) {
  return ejecutarConLock(function() {
    const idOrden = (datos.idOrden || "").toString();
    const monto = parseFloat(datos.monto);
    const fecha = datos.fecha || new Date().toISOString().slice(0, 10);
    const metodo = (datos.metodo || "Transferencia").toString();
    const nota = (datos.nota || "").toString();

    if (!idOrden) return { exito: false, mensaje: "Falta indicar la orden de compra." };
    if (isNaN(monto) || monto <= 0) return { exito: false, mensaje: "El monto del hito debe ser mayor a cero." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen órdenes registradas." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idOrden) continue;

      let items = [];
      try { items = JSON.parse(db[i][5] || "[]"); } catch(e) {}
      let hitos = [];
      try { hitos = JSON.parse(db[i][10] || "[]"); } catch(e) {}

      const totalOrden = calcularTotalAPagarOrden(items);
      const totalYaPagado = hitos.reduce((acc, h) => acc + (parseFloat(h.monto) || 0), 0);

      if (totalYaPagado + monto > totalOrden + 0.01) {
        const disponible = Math.max(0, totalOrden - totalYaPagado);
        return { exito: false, mensaje: "⚠️ El monto excede el total de la orden. Disponible para pagar: $" + disponible.toFixed(2) + "." };
      }

      hitos.push({
        id: "HITO-" + new Date().getTime(),
        monto: monto,
        fecha: fecha,
        metodo: metodo,
        nota: nota,
        registradoPor: Session.getActiveUser().getEmail() || ""
      });

      hoja.getRange(i + 1, 11).setValue(JSON.stringify(hitos));

      const nuevoTotalPagado = totalYaPagado + monto;
      const nuevoEstadoPago = nuevoTotalPagado >= totalOrden - 0.01 ? "Pagada" : (nuevoTotalPagado > 0 ? "Parcial" : "Pendiente");
      hoja.getRange(i + 1, 10).setValue(nuevoEstadoPago);

      verificarYCerrarOrdenSiCorresponde(hoja, i + 1, db[i]);

      registrarAuditoria("Compras", "Editar", "Hito de pago registrado en " + idOrden + ": $" + monto.toFixed(2) + " (" + metodo + ")");
      return { exito: true, mensaje: "¡Hito de pago de $" + monto.toFixed(2) + " registrado con éxito!", estadoPago: nuevoEstadoPago };
    }
    return { exito: false, mensaje: "No se encontró la orden." };
  });
}

function eliminarHitoPagoOrdenServidor(idOrden, idHito) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordenes_Compra");
    if (!hoja) return { exito: false, mensaje: "No existen órdenes registradas." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idOrden.toString()) continue;

      let items = [];
      try { items = JSON.parse(db[i][5] || "[]"); } catch(e) {}
      let hitos = [];
      try { hitos = JSON.parse(db[i][10] || "[]"); } catch(e) {}

      const hitosFiltrados = hitos.filter(h => h.id !== idHito);
      if (hitosFiltrados.length === hitos.length) {
        return { exito: false, mensaje: "No se encontró el hito de pago." };
      }

      hoja.getRange(i + 1, 11).setValue(JSON.stringify(hitosFiltrados));

      const totalOrden = calcularTotalAPagarOrden(items);
      const totalPagado = hitosFiltrados.reduce((acc, h) => acc + (parseFloat(h.monto) || 0), 0);
      const nuevoEstadoPago = totalPagado >= totalOrden - 0.01 ? "Pagada" : (totalPagado > 0 ? "Parcial" : "Pendiente");
      hoja.getRange(i + 1, 10).setValue(nuevoEstadoPago);

      registrarAuditoria("Compras", "Eliminar", "Hito de pago eliminado de " + idOrden);
      return { exito: true, mensaje: "¡Hito de pago eliminado con éxito!", estadoPago: nuevoEstadoPago };
    }
    return { exito: false, mensaje: "No se encontró la orden." };
  });
}

/**
 * Helper interno: calcula el total a pagar de una orden a partir de sus
 * items (mismo calculo usado en obtenerOrdenesCompraServidor, extraido aqui
 * para reutilizarlo al validar hitos de pago sin duplicar la formula).
 */
/**
 * Helper interno: normaliza una fecha (que puede llegar como objeto Date
 * real de Google Sheets, o como texto vacio) a un string simple "YYYY-MM-DD",
 * compatible con <input type="date"> y facil de parsear en el cliente. Evita
 * el problema de un Date object serializado a texto largo (ej. "Tue Jul 14
 * 2026 00:00:00 GMT-0600...") que rompe al concatenarlo con horas despues.
 */
function formatearFechaSimple(valor) {
  if (!valor) return "";
  try {
    const fecha = (valor instanceof Date) ? valor : new Date(valor);
    if (isNaN(fecha.getTime())) return "";
    const anio = fecha.getFullYear();
    const mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
    const dia = fecha.getDate().toString().padStart(2, "0");
    return anio + "-" + mes + "-" + dia;
  } catch (e) {
    return "";
  }
}

function calcularTotalAPagarOrden(items) {
  return (items || []).reduce((acc, it) => {
    const subtotalItem = (parseFloat(it.cantidadOrdenada) || 0) * (parseFloat(it.costoUnitario) || 0);
    const totalCargosItem = (it.cargos || []).reduce((accCargo, cargo) => {
      const montoCargo = subtotalItem * ((parseFloat(cargo.porcentaje) || 0) / 100);
      return accCargo + (cargo.tipo === 'Resta' ? -montoCargo : montoCargo);
    }, 0);
    return acc + subtotalItem + totalCargosItem;
  }, 0);
}

// ==========================================================================
//   REPORTES PRO DE COMPRAS (exportacion a Excel)
// ==========================================================================

/**
 * Listado completo de Ordenes de Compra para exportar a Excel: una fila por
 * orden con sus totales, estados, y hitos de pago resumidos.
 */
function obtenerDatosExportacionOrdenesServidor() {
  try {
    const ordenes = obtenerOrdenesCompraServidor();
    const filas = [["ID Orden", "Fecha", "Proveedor", "Origen", "Bodega Destino", "Estado Recepción", "Estado Pago", "Total a Pagar", "Total Pagado", "Saldo Pendiente", "N° de Ítems", "Requisición(es) Origen", "Notas"]];

    ordenes.forEach(o => {
      filas.push([
        o.id,
        o.fecha ? new Date(o.fecha).toLocaleString() : "",
        o.proveedor,
        o.origenTipo,
        o.bodegaDestino || "",
        o.estado,
        o.estadoPago,
        o.totalAPagar,
        o.totalPagado,
        o.saldoPendiente,
        (o.items || []).length,
        o.idRequisicion || "",
        o.notas || ""
      ]);
    });
    return filas;
  } catch (e) {
    console.error("Error al exportar órdenes de compra: " + e.message);
    return [["Error al generar el archivo"]];
  }
}

/**
 * Varianza de precio por producto: para cada SKU comprado, muestra el costo
 * unitario minimo, maximo, promedio, y cuanto ha variado en el tiempo entre
 * distintas ordenes de compra. Util para detectar si un proveedor empezo a
 * cobrar mas caro, o si hay diferencias grandes entre compras del mismo item.
 */
function obtenerDatosExportacionVarianzaPrecioServidor() {
  try {
    const ordenes = obtenerOrdenesCompraServidor().filter(o => o.estado !== 'Cancelada');
    const porSku = {};

    ordenes.forEach(o => {
      (o.items || []).forEach(it => {
        const costo = parseFloat(it.costoUnitario) || 0;
        if (costo <= 0) return;
        if (!porSku[it.sku]) porSku[it.sku] = { descripcion: it.descripcion, compras: [] };
        porSku[it.sku].compras.push({ costo: costo, fecha: o.fecha, proveedor: o.proveedor, idOrden: o.id });
      });
    });

    const filas = [["SKU", "Descripción", "N° de Compras", "Costo Mínimo", "Costo Máximo", "Costo Promedio", "Variación ($)", "Variación (%)", "Última Compra", "Proveedor Última Compra", "Costo Última Compra"]];

    Object.keys(porSku).forEach(sku => {
      const info = porSku[sku];
      const costos = info.compras.map(c => c.costo);
      const minimo = Math.min(...costos);
      const maximo = Math.max(...costos);
      const promedio = costos.reduce((a, b) => a + b, 0) / costos.length;
      const variacionAbs = maximo - minimo;
      const variacionPct = minimo > 0 ? (variacionAbs / minimo) * 100 : 0;

      const ordenadosPorFecha = [...info.compras].sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
      const ultima = ordenadosPorFecha[0];

      filas.push([
        sku, info.descripcion, info.compras.length,
        minimo, maximo, Math.round(promedio * 100) / 100,
        Math.round(variacionAbs * 100) / 100, Math.round(variacionPct * 100) / 100,
        ultima.fecha ? new Date(ultima.fecha).toLocaleDateString() : "",
        ultima.proveedor, ultima.costo
      ]);
    });

    return filas;
  } catch (e) {
    console.error("Error al exportar varianza de precio: " + e.message);
    return [["Error al generar el archivo"]];
  }
}

/**
 * Comparativa entre proveedores: para cada proveedor, totales comprados,
 * numero de ordenes, costo promedio por item, y -- cuando dos o mas
 * proveedores han vendido el MISMO sku -- una comparacion directa de cual
 * ofrece mejor precio para ese item especifico.
 */
function obtenerDatosExportacionComparativaProveedoresServidor() {
  try {
    const ordenes = obtenerOrdenesCompraServidor().filter(o => o.estado !== 'Cancelada');

    // Hoja 1: resumen general por proveedor
    const resumenPorProveedor = {};
    ordenes.forEach(o => {
      if (!resumenPorProveedor[o.proveedor]) {
        resumenPorProveedor[o.proveedor] = { ordenes: 0, totalComprado: 0, items: 0 };
      }
      resumenPorProveedor[o.proveedor].ordenes++;
      resumenPorProveedor[o.proveedor].totalComprado += o.totalAPagar || 0;
      resumenPorProveedor[o.proveedor].items += (o.items || []).length;
    });

    const filasResumen = [["Proveedor", "N° de Órdenes", "Total Comprado", "N° de Ítems Comprados", "Promedio por Orden"]];
    Object.keys(resumenPorProveedor).forEach(prov => {
      const r = resumenPorProveedor[prov];
      filasResumen.push([prov, r.ordenes, Math.round(r.totalComprado * 100) / 100, r.items, Math.round((r.totalComprado / r.ordenes) * 100) / 100]);
    });

    // Hoja 2: comparativa directa por SKU vendido por mas de un proveedor
    const porSkuYProveedor = {};
    ordenes.forEach(o => {
      (o.items || []).forEach(it => {
        const costo = parseFloat(it.costoUnitario) || 0;
        if (costo <= 0) return;
        const clave = it.sku;
        if (!porSkuYProveedor[clave]) porSkuYProveedor[clave] = { descripcion: it.descripcion, proveedores: {} };
        if (!porSkuYProveedor[clave].proveedores[o.proveedor]) porSkuYProveedor[clave].proveedores[o.proveedor] = [];
        porSkuYProveedor[clave].proveedores[o.proveedor].push(costo);
      });
    });

    const filasComparativa = [["SKU", "Descripción", "Proveedor", "Costo Promedio", "N° de Compras", "¿Mejor Precio?"]];
    Object.keys(porSkuYProveedor).forEach(sku => {
      const info = porSkuYProveedor[sku];
      const nombresProveedores = Object.keys(info.proveedores);
      if (nombresProveedores.length < 2) return; // solo interesa comparar si hay 2+ proveedores del mismo item

      const promediosPorProveedor = nombresProveedores.map(prov => {
        const costos = info.proveedores[prov];
        return { proveedor: prov, promedio: costos.reduce((a, b) => a + b, 0) / costos.length, compras: costos.length };
      });

      const mejorPrecio = Math.min(...promediosPorProveedor.map(p => p.promedio));

      promediosPorProveedor.forEach(p => {
        filasComparativa.push([
          sku, info.descripcion, p.proveedor,
          Math.round(p.promedio * 100) / 100, p.compras,
          p.promedio === mejorPrecio ? "Sí" : "No"
        ]);
      });
    });

    return { resumen: filasResumen, comparativa: filasComparativa };
  } catch (e) {
    console.error("Error al exportar comparativa de proveedores: " + e.message);
    return { resumen: [["Error al generar el archivo"]], comparativa: [] };
  }
}

/**
 * Helper interno: si la orden ya esta 100% recibida Y 100% pagada, la marca
 * con estado final "Cerrada" -- la orden no se considera completamente
 * cerrada hasta que ambas condiciones se cumplen, sin importar cual de las
 * dos se complete primero.
 */
function verificarYCerrarOrdenSiCorresponde(hoja, numeroFila, filaOrden) {
  try {
    let items = [];
    try { items = JSON.parse(filaOrden[5] || "[]"); } catch(e) {}
    const estadoRecepcion = (filaOrden[4] || "").toString();
    const estadoPagoActual = hoja.getRange(numeroFila, 10).getValue().toString();

    const recepcionCompleta = estadoRecepcion === "Recibida Completa";
    const pagoCompleto = estadoPagoActual === "Pagada";

    if (recepcionCompleta && pagoCompleto) {
      hoja.getRange(numeroFila, 5).setValue("Cerrada");
    }
  } catch (e) {
    console.error("Error al verificar cierre de orden: " + e.message);
  }
}



/**
 * Registra la recepcion (parcial o total) de una Orden de Compra. Cada
 * item recibido genera un movimiento real de "Entrada / Compra Reventa"
 * en Inventario (actualiza stock + costo promedio), igual que la entrada
 * manual que existia en el modulo de Inventarios.
 * @param {Object} datos { idOrden, bodega, items: [{sku, cantidadRecibida}] }
 */
function registrarRecepcionCompraServidor(datos) {
  return ejecutarConLock(function() {
    const idOrden = datos.idOrden;
    const itemsRecibidos = datos.items || [];
    const recibidoPor = (datos.recibidoPor || "").toString().trim();
    const firmaDigital = (datos.firmaDigital || "").toString(); // placeholder PRO -- la captura real de firma se construira en una sesion movil dedicada

    if (!idOrden) return { exito: false, mensaje: "Falta indicar la orden de compra." };
    if (itemsRecibidos.length === 0) return { exito: false, mensaje: "Indica al menos un ítem recibido." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaOrdenes = ss.getSheetByName("Ordenes_Compra");
    if (!hojaOrdenes) return { exito: false, mensaje: "No existe la orden de compra." };

    const db = hojaOrdenes.getDataRange().getValues();
    let filaOrden = -1;
    let itemsOrden = [];
    let proveedorOrden = "";
    let tipoDestino = "Bodega";
    let bodegaDestino = "Sin Bodega";
    let entregaDirectaA = "";

    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() === idOrden.toString()) {
        filaOrden = i + 1;
        try { itemsOrden = JSON.parse(db[i][5] || "[]"); } catch(e) {}
        proveedorOrden = (db[i][3] || "").toString();
        bodegaDestino = (db[i][8] || "Sin Bodega").toString();
        tipoDestino = (db[i][11] || "Bodega").toString();
        entregaDirectaA = (db[i][12] || "").toString();
        break;
      }
    }

    if (filaOrden === -1) return { exito: false, mensaje: "No se encontró la orden de compra." };

    const esEntregaDirecta = tipoDestino === "Entrega Directa";
    if (esEntregaDirecta && !recibidoPor) {
      return { exito: false, mensaje: "Indica el nombre de quién recibió la entrega." };
    }

    // Validar que no se reciba mas de lo pendiente, y actualizar cantidadRecibida acumulada
    for (let i = 0; i < itemsRecibidos.length; i++) {
      const itemRecibido = itemsRecibidos[i];
      const itemOrden = itemsOrden.find(it => it.sku === itemRecibido.sku);
      if (!itemOrden) continue;

      const pendiente = (itemOrden.cantidadOrdenada || 0) - (itemOrden.cantidadRecibida || 0);
      if ((itemRecibido.cantidadRecibida || 0) > pendiente) {
        return { exito: false, mensaje: `⚠️ No puedes recibir más de lo pendiente para "${itemOrden.descripcion || itemOrden.sku}". Pendiente: ${pendiente}.` };
      }
    }

    // Se crea/obtiene la hoja de Recepciones ANTES de generar el correlativo,
    // para poder contar correctamente cuantos recibos WH-IN ya existen.
    const ss2 = SpreadsheetApp.getActiveSpreadsheet();
    let hojaRecepciones = ss2.getSheetByName("Recepciones");
    if (!hojaRecepciones) {
      hojaRecepciones = ss2.insertSheet("Recepciones");
      hojaRecepciones.appendRow(["ID", "Fecha", "ID_Orden", "Items (JSON)", "Usuario", "Tipo Destino", "Recibido Por", "Entrega Directa A", "Firma Digital", "Tipo Movimiento", "Bodega"]);
    }
    const idRecepcion = generarCorrelativoRecepcionServidor(hojaRecepciones, "WH-IN");

    if (esEntregaDirecta) {
      // ENTREGA DIRECTA: no hay bodega que controlar, asi que NO se generan
      // movimientos de inventario (no suma al stock vendible). Sin embargo,
      // el costo de esta compra SI debe reflejarse en el costo promedio
      // global del producto -- igual al patron ya usado en "Compra Consumo
      // Interno" desde Inventarios -- para que el promedio refleje todo lo
      // pagado por ese SKU, sin importar el destino de la compra.
      const ss2 = SpreadsheetApp.getActiveSpreadsheet();
      let hojaHistCostos = ss2.getSheetByName("Historial_Costos");
      if (!hojaHistCostos) {
        hojaHistCostos = ss2.insertSheet("Historial_Costos");
        hojaHistCostos.appendRow(["Fecha", "SKU", "Origen", "Cantidad", "Costo Unitario", "Nota"]);
      }

      itemsRecibidos.forEach(itemRecibido => {
        const cantidad = parseFloat(itemRecibido.cantidadRecibida) || 0;
        if (cantidad <= 0) return;
        const itemOrden = itemsOrden.find(it => it.sku === itemRecibido.sku);
        if (itemOrden) itemOrden.cantidadRecibida = (itemOrden.cantidadRecibida || 0) + cantidad;

        const costoUnitario = itemOrden ? (parseFloat(itemOrden.costoUnitario) || 0) : 0;
        if (costoUnitario > 0) {
          actualizarCostoPromedioServidor(itemRecibido.sku, cantidad, costoUnitario, "Compra Consumo Interno");
          hojaHistCostos.appendRow([
            new Date(), itemRecibido.sku, "Compra Consumo Interno (Entrega Directa)", cantidad, costoUnitario,
            "Proveedor: " + proveedorOrden + " | Orden: " + idOrden + " | Recepción: " + idRecepcion + " | Entregado a: " + entregaDirectaA
          ]);
        }
      });
    } else {
      // BODEGA: registra movimientos de entrada reales en inventario y
      // actualiza el costo promedio, igual que antes.
      const hojaMov = obtenerOCrearHojaMovimientos(ss);

      itemsRecibidos.forEach(itemRecibido => {
        const cantidad = parseFloat(itemRecibido.cantidadRecibida) || 0;
        if (cantidad <= 0) return;

        const itemOrden = itemsOrden.find(it => it.sku === itemRecibido.sku);
        const costoUnitario = itemOrden ? (parseFloat(itemOrden.costoUnitario) || 0) : 0;

        hojaMov.appendRow([
          "MOV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000),
          new Date(), itemRecibido.sku, "Entrada", "Compra Reventa",
          cantidad, bodegaDestino, idOrden,
          "Proveedor: " + proveedorOrden + (costoUnitario ? (" | Costo unitario: $" + costoUnitario) : "") + " | Recepción: " + idRecepcion,
          Session.getActiveUser().getEmail() || ""
        ]);

        if (costoUnitario > 0) actualizarCostoPromedioServidor(itemRecibido.sku, cantidad, costoUnitario, "Compra Reventa");

        if (itemOrden) itemOrden.cantidadRecibida = (itemOrden.cantidadRecibida || 0) + cantidad;
      });
    }

    // Determinar nuevo estado de la orden: completa si todos los items llegaron al 100%
    const todoCompleto = itemsOrden.every(it => (it.cantidadRecibida || 0) >= (it.cantidadOrdenada || 0));
    const algoRecibido = itemsOrden.some(it => (it.cantidadRecibida || 0) > 0);
    const nuevoEstado = todoCompleto ? "Recibida Completa" : (algoRecibido ? "Recibida Parcial" : "Pendiente de Recepción");

    hojaOrdenes.getRange(filaOrden, 5).setValue(nuevoEstado);
    hojaOrdenes.getRange(filaOrden, 6).setValue(JSON.stringify(itemsOrden));

    if (nuevoEstado === "Recibida Completa") {
      const filaActualizada = hojaOrdenes.getRange(filaOrden, 1, 1, 16).getValues()[0];
      verificarYCerrarOrdenSiCorresponde(hojaOrdenes, filaOrden, filaActualizada);
    }

    // Propagar "Recibido" al estadoItem dentro de las Requisiciones de origen,
    // solo para los items cuya cantidad recibida ya alcanzo lo ordenado.
    const idsReqOrigen = (db[filaOrden - 1][2] || "").toString().split(",").map(s => s.trim()).filter(Boolean);
    if (idsReqOrigen.length > 0) {
      const hojaReq = ss.getSheetByName("Requisiciones");
      if (hojaReq) {
        const dbReq = hojaReq.getDataRange().getValues();
        const skusCompletados = itemsOrden.filter(it => (it.cantidadRecibida || 0) >= (it.cantidadOrdenada || 0)).map(it => it.sku);

        for (let i = 1; i < dbReq.length; i++) {
          const idReq = dbReq[i][0].toString();
          if (!idsReqOrigen.includes(idReq)) continue;

          let itemsReq = [];
          try { itemsReq = JSON.parse(dbReq[i][5] || "[]"); } catch(e) {}

          let cambio = false;
          itemsReq.forEach(itemReq => {
            if (itemReq.idOrdenAsociada === idOrden && skusCompletados.includes(itemReq.sku)) {
              itemReq.estadoItem = "Recibido";
              cambio = true;
            }
          });

          if (cambio) hojaReq.getRange(i + 1, 6).setValue(JSON.stringify(itemsReq));
        }
      }
    }

    // Registrar la recepcion en su propia hoja para trazabilidad. Para
    // Entrega Directa, ademas se guarda quien recibio y a quien iba dirigida.
    // Se guarda el NOMBRE del usuario (no el correo crudo) para que se vea
    // mejor en el recibo, y la bodega real de destino (vacia si fue Entrega
    // Directa, ya que ahi no aplica ninguna bodega).
    hojaRecepciones.appendRow([
      idRecepcion, new Date(), idOrden, JSON.stringify(itemsRecibidos), resolverNombreUsuarioPorCorreoServidor(Session.getActiveUser().getEmail()),
      tipoDestino, recibidoPor, entregaDirectaA, firmaDigital, "Entrada", esEntregaDirecta ? "" : bodegaDestino
    ]);

    registrarAuditoria("Recepción", "Editar", "Recepción registrada para Orden: " + idOrden + (esEntregaDirecta ? (" (Entrega Directa, recibido por: " + recibidoPor + ")") : ""));
    return { exito: true, mensaje: "¡Recepción registrada con éxito! Estado de la orden: " + nuevoEstado, estado: nuevoEstado };
  });
}

/**
 * Devuelve TODAS las recepciones ya registradas (cada confirmacion individual,
 * no las ordenes en si), mas recientes primero. Se usa para el historial de
 * recibos tipo WH-IN/WH-OUT en el modulo de Recepcion.
 */
function obtenerRecepcionesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Recepciones");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const recepciones = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][3] || "[]"); } catch(e) {}

      recepciones.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        idOrden: datos[i][2] ? datos[i][2].toString() : "",
        items: items,
        usuario: datos[i][4] ? datos[i][4].toString() : "",
        tipoDestino: datos[i][5] ? datos[i][5].toString() : "Bodega",
        recibidoPor: datos[i][6] ? datos[i][6].toString() : "",
        entregaDirectaA: datos[i][7] ? datos[i][7].toString() : "",
        firmaDigital: datos[i][8] ? datos[i][8].toString() : "",
        tipoMovimiento: datos[i][9] ? datos[i][9].toString() : "Entrada",
        bodega: datos[i][10] ? datos[i][10].toString() : ""
      });
    }
    recepciones.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return recepciones;
  } catch (e) {
    console.error("Error al obtener recepciones: " + e.message);
    return [];
  }
}

/**
 * Construye el detalle COMPLETO de un recibo de recepcion (tipo WH-IN/WH-OUT):
 * cruza la recepcion puntual con su Orden de Compra de origen y, si aplica,
 * con la(s) Requisicion(es) que dieron pie a esa orden -- todo en un solo
 * objeto listo para pintar en el modal de "Ver Recibo" o generar su PDF.
 */
function obtenerDetalleRecepcionServidor(idRecepcion) {
  try {
    const recepciones = obtenerRecepcionesServidor();
    const recepcion = recepciones.find(r => r.id === idRecepcion);
    if (!recepcion) return { exito: false, mensaje: "No se encontró la recepción." };

    const ordenes = obtenerOrdenesCompraServidor();
    const orden = ordenes.find(o => o.id === recepcion.idOrden);

    let requisiciones = [];
    if (orden && orden.idRequisicion) {
      const idsReq = orden.idRequisicion.split(",").map(s => s.trim()).filter(Boolean);
      const todasLasReq = obtenerRequisicionesServidor();
      requisiciones = todasLasReq.filter(r => idsReq.includes(r.id));
    }

    // Enriquece cada item recibido con su descripcion completa. Para
    // entradas (con orden de origen), la descripcion vive en la orden. Para
    // salidas (sin orden), se resuelve directamente desde el catalogo de
    // productos a partir del SKU.
    let mapaDescripcionPorSku = {};
    if (!orden) {
      const catalogo = obtenerCatalogoParaRequisicionesServidor();
      catalogo.forEach(p => (p.variantes || []).forEach(v => {
        mapaDescripcionPorSku[v.sku] = p.nombre + ' - ' + (v.etiqueta || 'Estándar');
      }));
    }

    const itemsConDescripcion = recepcion.items.map(it => {
      const itemOrden = orden ? (orden.items || []).find(io => io.sku === it.sku) : null;
      return Object.assign({}, it, {
        descripcion: itemOrden ? itemOrden.descripcion : (mapaDescripcionPorSku[it.sku] || it.sku),
        costoUnitario: itemOrden ? (parseFloat(itemOrden.costoUnitario) || 0) : 0,
        idRequisicionOrigen: itemOrden ? (itemOrden.idRequisicionOrigen || "") : ""
      });
    });

    return {
      exito: true,
      recepcion: Object.assign({}, recepcion, { items: itemsConDescripcion }),
      orden: orden || null,
      requisiciones: requisiciones
    };
  } catch (e) {
    return { exito: false, mensaje: "Error al construir el detalle del recibo: " + e.message };
  }
}

// ==========================================================================
//   MÓDULO: VENTAS — HISTORIAL Y GESTIÓN DE COTIZACIONES
// ==========================================================================

/**
 * Devuelve todas las cotizaciones guardadas, más recientes primero.
 * Se usa para el panel "Cotizaciones" en ModVentas.
 */
function obtenerCotizacionesServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cotizaciones");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const cotizaciones = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][6] || "[]"); } catch(e) {}
      cotizaciones.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        clienteNombre: datos[i][2] ? datos[i][2].toString() : "",
        clienteDuiNit: datos[i][3] ? datos[i][3].toString() : "",
        tipoCliente: datos[i][4] ? datos[i][4].toString() : "",
        correo: datos[i][5] ? datos[i][5].toString() : "",
        items: items,
        subtotal: parseFloat(datos[i][7]) || 0,
        iva: parseFloat(datos[i][8]) || 0,
        total: parseFloat(datos[i][9]) || 0,
        estado: datos[i][10] ? datos[i][10].toString() : "Pendiente de Aprobación",
        notas: datos[i][11] ? datos[i][11].toString() : "",
        telefono: (() => {
          // Busca el telefono del cliente en la hoja Clientes
          try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const hCli = ss.getSheetByName("Clientes");
            if (!hCli) return "";
            const dbCli = hCli.getDataRange().getValues();
            const duiCot = datos[i][3] ? datos[i][3].toString() : "";
            for (let j = 1; j < dbCli.length; j++) {
              if ((dbCli[j][1] || "").toString() === duiCot) return (dbCli[j][3] || "").toString();
            }
          } catch(e) {}
          return "";
        })()
      });
    }
    cotizaciones.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return cotizaciones;
  } catch (e) {
    console.error("Error al obtener cotizaciones: " + e.message);
    return [];
  }
}

/**
 * Convierte una cotizacion existente en un Pedido a Produccion o en una
 * Venta Directa, evitando tener que volver a capturar todos los datos.
 * Cambia el estado de la cotizacion a "Convertida" y procesa la accion.
 */
function convertirCotizacionServidor(idCotizacion, tipoConversion, formaPago) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cotizaciones");
    if (!hoja) return { exito: false, mensaje: "No existen cotizaciones registradas." };

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString() !== idCotizacion.toString()) continue;

      const estado = (datos[i][10] || "").toString();
      if (estado === "Convertida") {
        return { exito: false, mensaje: "Esta cotización ya fue convertida anteriormente." };
      }

      let items = [];
      try { items = JSON.parse(datos[i][6] || "[]"); } catch(e) {}

      const payload = {
        cliente: {
          nombre: datos[i][2],
          duiNit: datos[i][3],
          tipoCliente: datos[i][4],
          correo: datos[i][5]
        },
        artículos: items,
        financiero: {
          subtotal: parseFloat(datos[i][7]) || 0,
          iva: parseFloat(datos[i][8]) || 0,
          total: parseFloat(datos[i][9]) || 0
        },
        formaPago: formaPago || "",
        notas: "Convertida desde cotización " + idCotizacion + (datos[i][11] ? " · " + datos[i][11] : "")
      };

      let resultado;
      if (tipoConversion === "VENTA_DIRECTA") {
        if (!formaPago) return { exito: false, mensaje: "Indica la forma de pago para convertir en venta." };
        resultado = procesarVentaDirectaServidor(payload);
      } else {
        resultado = procesarPedidoProduccionServidor(payload);
      }

      if (resultado.exito) {
        // Marcar la cotizacion como convertida
        hoja.getRange(i + 1, 11).setValue("Convertida");
        registrarAuditoria("Ventas", "Convertir", "Cotización " + idCotizacion + " convertida a " + tipoConversion);
      }

      return resultado;
    }
    return { exito: false, mensaje: "No se encontró la cotización." };
  });
}

/**
 * Devuelve el historial de ventas directas, más recientes primero.
 */
function obtenerVentasDirectasServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ventas_Directas");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const ventas = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][5] || "[]"); } catch(e) {}
      ventas.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        clienteNombre: datos[i][2] ? datos[i][2].toString() : "",
        clienteDuiNit: datos[i][3] ? datos[i][3].toString() : "",
        tipoCliente: datos[i][4] ? datos[i][4].toString() : "",
        items: items,
        subtotal: parseFloat(datos[i][6]) || 0,
        iva: parseFloat(datos[i][7]) || 0,
        total: parseFloat(datos[i][8]) || 0,
        formaPago: datos[i][9] ? datos[i][9].toString() : "",
        estado: datos[i][10] ? datos[i][10].toString() : "",
        notas: datos[i][11] ? datos[i][11].toString() : ""
      });
    }
    ventas.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return ventas;
  } catch (e) {
    console.error("Error al obtener ventas directas: " + e.message);
    return [];
  }
}

// ==========================================================================
//   MÓDULO: LISTAS DE PRECIOS
// ==========================================================================

/**
 * Devuelve todas las listas de precios definidas.
 * Estructura de la hoja Listas_Precios:
 *   ID_Lista | Nombre | Descripcion | Items_JSON | Clientes_JSON | Fecha
 * Items_JSON: [{ sku, descripcion, precioBase, precioLista }]
 * Clientes_JSON: ["duiNit1", "duiNit2", ...]
 */
function obtenerListasPreciosServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Listas_Precios");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const listas = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [], clientes = [];
      try { items = JSON.parse(datos[i][3] || "[]"); } catch(e) {}
      try { clientes = JSON.parse(datos[i][4] || "[]"); } catch(e) {}
      listas.push({
        id: datos[i][0].toString(),
        nombre: datos[i][1] ? datos[i][1].toString() : "",
        descripcion: datos[i][2] ? datos[i][2].toString() : "",
        items: items,
        clientes: clientes,
        fecha: datos[i][5] ? datos[i][5].toString() : "",
        esPredeterminada: datos[i][6] === true || datos[i][6] === "TRUE"
      });
    }
    return listas;
  } catch (e) {
    console.error("Error al obtener listas de precios: " + e.message);
    return [];
  }
}

/**
 * Guarda o actualiza una lista de precios.
 * Si datos.id existe, actualiza; si no, crea una nueva.
 */
function guardarListaPreciosServidor(datos) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Listas_Precios");
    if (!hoja) {
      hoja = ss.insertSheet("Listas_Precios");
      hoja.appendRow(["ID_Lista", "Nombre", "Descripcion", "Items_JSON", "Clientes_JSON", "Fecha", "Es_Predeterminada"]);
    }

    const id = datos.id || ("LP-" + new Date().getTime());

    // Si esta lista se marca como predeterminada, quita ese flag de todas las demás
    if (datos.esPredeterminada) {
      const db = hoja.getDataRange().getValues();
      for (let i = 1; i < db.length; i++) {
        if (db[i][0].toString() !== id) {
          hoja.getRange(i + 1, 7).setValue(false);
        }
      }
    }

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() === id) {
        hoja.getRange(i + 1, 1, 1, 7).setValues([[
          id, datos.nombre || "", datos.descripcion || "",
          JSON.stringify(datos.items || []),
          JSON.stringify(datos.clientes || []),
          new Date(),
          datos.esPredeterminada === true
        ]]);
        registrarAuditoria("Ventas", "Editar", "Lista de precios actualizada: " + datos.nombre);
        return { exito: true, mensaje: "Lista de precios actualizada.", id: id };
      }
    }

    hoja.appendRow([
      id, datos.nombre || "", datos.descripcion || "",
      JSON.stringify(datos.items || []),
      JSON.stringify(datos.clientes || []),
      new Date(),
      datos.esPredeterminada === true
    ]);
    registrarAuditoria("Ventas", "Crear", "Lista de precios creada: " + datos.nombre);
    return { exito: true, mensaje: "Lista de precios creada.", id: id };
  });
}

/**
 * Elimina una lista de precios por ID.
 */
function eliminarListaPreciosServidor(idLista) {
  return ejecutarConLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Listas_Precios");
    if (!hoja) return { exito: false, mensaje: "No existen listas de precios." };
    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() === idLista) {
        hoja.deleteRow(i + 1);
        return { exito: true, mensaje: "Lista eliminada." };
      }
    }
    return { exito: false, mensaje: "No se encontró la lista." };
  });
}

/**
 * Resuelve la lista de precios aplicable a un cliente con esta prioridad:
 * 1. Si el cliente tiene una lista especifica asignada → esa lista para sus
 *    productos incluidos, y la predeterminada como fallback para el resto
 * 2. Si no tiene lista especifica → la lista predeterminada
 * 3. Si no hay nada → null (usa precio base del catalogo)
 *
 * Devuelve un objeto { items, nombre, combinada } donde items es la union
 * de ambas listas con la especifica teniendo prioridad producto por producto.
 */
function obtenerListaPreciosPorClienteServidor(duiNit) {
  try {
    const listas = obtenerListasPreciosServidor();
    const duiLimpio = (duiNit || "").toString().trim();

    // Buscar lista especifica del cliente
    const listaEspecifica = listas.find(l =>
      !l.esPredeterminada && (l.clientes || []).some(c => c.toString().trim() === duiLimpio)
    );

    // Buscar lista predeterminada
    const listaPredeterminada = listas.find(l => l.esPredeterminada);

    if (!listaEspecifica && !listaPredeterminada) return null;

    if (listaEspecifica && !listaPredeterminada) return listaEspecifica;

    if (!listaEspecifica && listaPredeterminada) return listaPredeterminada;

    // Tiene ambas: combinar. La especifica tiene prioridad por SKU/descripcion.
    const itemsCombinados = [...listaPredeterminada.items];
    (listaEspecifica.items || []).forEach(itemEsp => {
      const idx = itemsCombinados.findIndex(it =>
        it.sku === itemEsp.sku || it.descripcion === itemEsp.descripcion
      );
      if (idx >= 0) {
        itemsCombinados[idx] = itemEsp; // reemplaza con precio especial
      } else {
        itemsCombinados.push(itemEsp); // agrega el producto especial
      }
    });

    return {
      id: listaEspecifica.id,
      nombre: listaEspecifica.nombre + ' + ' + listaPredeterminada.nombre,
      items: itemsCombinados,
      clientes: listaEspecifica.clientes,
      combinada: true
    };
  } catch (e) {
    console.error("Error al obtener lista de precios por cliente: " + e.message);
    return null;
  }
}

/**
 * Envia una cotizacion guardada por correo al cliente (solo PRO).
 * Genera un resumen en texto con los items y totales.
 */
function enviarCotizacionPorCorreoServidor(idCotizacion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Cotizaciones");
    if (!hoja) return { exito: false, mensaje: "No existen cotizaciones." };

    const db = hoja.getDataRange().getValues();
    for (let i = 1; i < db.length; i++) {
      if (db[i][0].toString() !== idCotizacion.toString()) continue;

      const correo = db[i][5] ? db[i][5].toString() : "";
      const cliente = db[i][2] ? db[i][2].toString() : "";
      if (!correo) return { exito: false, mensaje: "El cliente no tiene correo registrado." };

      let items = [];
      try { items = JSON.parse(db[i][6] || "[]"); } catch(e) {}
      const total = parseFloat(db[i][9]) || 0;
      const notas = db[i][11] ? db[i][11].toString() : "";

      const lineasItems = items.map(it =>
        "• " + it.desc + " — Cant: " + it.cant + " @ $" + parseFloat(it.prec || 0).toFixed(2) + " = $" + parseFloat(it.total || 0).toFixed(2)
      ).join("\n");

      const empresa = obtenerConfiguracionEmpresaServidor() || {};
      const nombreEmpresa = empresa.nombreEmpresa || "Nuestra empresa";

      const asunto = "Cotización " + idCotizacion + " — " + nombreEmpresa;
      const cuerpo = "Estimado/a " + cliente + ",\n\n" +
        "Le compartimos el detalle de su cotización " + idCotizacion + ":\n\n" +
        lineasItems + "\n\n" +
        "Total: $" + total.toFixed(2) + "\n\n" +
        (notas ? "Notas: " + notas + "\n\n" : "") +
        "Quedamos a su disposición para cualquier consulta.\n\n" +
        "Atentamente,\n" + nombreEmpresa;

      MailApp.sendEmail(correo, asunto, cuerpo);
      registrarAuditoria("Ventas", "Enviar", "Cotización " + idCotizacion + " enviada a " + correo);
      return { exito: true, mensaje: "Cotización enviada a " + correo + "." };
    }
    return { exito: false, mensaje: "No se encontró la cotización." };
  } catch (e) {
    return { exito: false, mensaje: "Error al enviar: " + e.message };
  }
}

/**
 * Devuelve todos los pedidos a produccion para el historial unificado de ventas.
 */
function obtenerPedidosProduccionServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Pedidos_Produccion");
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];
    const pedidos = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let items = [];
      try { items = JSON.parse(datos[i][5] || "[]"); } catch(e) {}
      pedidos.push({
        id: datos[i][0].toString(),
        fecha: datos[i][1] ? datos[i][1].toString() : "",
        clienteNombre: datos[i][2] ? datos[i][2].toString() : "",
        clienteDuiNit: datos[i][3] ? datos[i][3].toString() : "",
        tipoCliente: datos[i][4] ? datos[i][4].toString() : "",
        items: items,
        subtotal: parseFloat(datos[i][6]) || 0,
        iva: parseFloat(datos[i][7]) || 0,
        total: parseFloat(datos[i][8]) || 0,
        estado: datos[i][9] ? datos[i][9].toString() : "En Producción",
        notas: datos[i][10] ? datos[i][10].toString() : ""
      });
    }
    pedidos.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return pedidos;
  } catch (e) {
    console.error("Error al obtener pedidos de producción: " + e.message);
    return [];
  }
}

// ==========================================================================
//   DOCUMENTOS FISCALES Y LIBRO DE INGRESOS
// ==========================================================================

/**
 * Genera un placeholder de documento fiscal en la hoja Documentos_Fiscales.
 * Estructura: ID_Doc | Fecha | Tipo_Doc | Tipo_Operacion | Cliente | DUI_NIT |
 *             Items_JSON | Subtotal | IVA | Total | Forma_Pago | Estado | Notas
 *
 * En el futuro este placeholder se usara para emitir factura electronica (DTE).
 * Por ahora sirve como registro de cada documento emitido y para generar el PDF.
 */
function generarDocumentoFiscalServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Documentos_Fiscales");
    if (!hoja) {
      hoja = ss.insertSheet("Documentos_Fiscales");
      hoja.appendRow([
        "ID_Documento", "ID_Referencia", "Fecha", "Tipo_Doc", "Tipo_Operacion",
        "Cliente", "DUI_NIT", "Items_JSON", "Subtotal", "IVA", "Total",
        "Forma_Pago", "Estado", "Notas"
      ]);
      hoja.setFrozenRows(1);
    }

    const anio = new Date().getFullYear();
    const db = hoja.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < db.length; i++) {
      const m = (db[i][0] || "").toString().match(/DOC-\d{4}-(\d+)/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1]));
    }
    const idDoc = "DOC-" + anio + "-" + String(maxNum + 1).padStart(4, "0");

    hoja.appendRow([
      idDoc,
      datos.id || "",
      new Date(),
      datos.tipo || "Consumidor_Final",
      datos.tipoOperacion || "",
      (datos.cliente || {}).nombre || "",
      (datos.cliente || {}).duiNit || "",
      JSON.stringify(datos.articulos || []),
      datos.financiero ? datos.financiero.subtotal : 0,
      datos.financiero ? datos.financiero.iva : 0,
      datos.financiero ? datos.financiero.total : 0,
      datos.formaPago || "",
      "Emitido",
      datos.notas || ""
    ]);

    return idDoc;
  } catch (e) {
    console.warn("No se pudo generar documento fiscal: " + e.message);
    return null;
  }
}

/**
 * Registra un ingreso o cuenta por cobrar en el libro de ingresos.
 * Estructura: ID | Fecha | Referencia | Tipo | Cliente | Monto | Forma_Pago | Estado
 * Sirve como base para el modulo contable futuro.
 */
function registrarIngresoServidor(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Libro_Ingresos");
    if (!hoja) {
      hoja = ss.insertSheet("Libro_Ingresos");
      hoja.appendRow([
        "ID_Ingreso", "Fecha", "Referencia", "Tipo", "Cliente",
        "Monto", "Forma_Pago", "Estado"
      ]);
      hoja.setFrozenRows(1);
    }

    const anio = new Date().getFullYear();
    const db = hoja.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < db.length; i++) {
      const m = (db[i][0] || "").toString().match(/ING-\d{4}-(\d+)/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1]));
    }
    const idIng = "ING-" + anio + "-" + String(maxNum + 1).padStart(4, "0");

    const esContado = datos.formaPago && datos.formaPago !== 'Pendiente de cobro';
    hoja.appendRow([
      idIng,
      new Date(),
      datos.referencia || "",
      datos.tipo || "",
      datos.cliente || "",
      datos.monto || 0,
      datos.formaPago || "",
      esContado ? "Cobrado" : "Pendiente"
    ]);

    return idIng;
  } catch (e) {
    console.warn("No se pudo registrar ingreso: " + e.message);
    return null;
  }
}

/**
 * Devuelve el historial unificado de ventas: ventas directas + pedidos a
 * produccion + cotizaciones convertidas. Ordenado por fecha descendente.
 * Este es el unico endpoint que necesita el historial de ventas en ModVentas.
 */
function obtenerHistorialVentasUnificadoServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultado = [];

    // Ventas Directas
    const hojaVD = ss.getSheetByName("Ventas_Directas");
    if (hojaVD) {
      const dbVD = hojaVD.getDataRange().getValues();
      for (let i = 1; i < dbVD.length; i++) {
        if (!dbVD[i][0]) continue;
        let items = [];
        try { items = JSON.parse(dbVD[i][5] || "[]"); } catch(e) {}
        resultado.push({
          id: dbVD[i][0].toString(),
          fecha: dbVD[i][1] ? dbVD[i][1].toString() : "",
          tipo: "Venta Directa",
          clienteNombre: dbVD[i][2] ? dbVD[i][2].toString() : "",
          clienteDuiNit: dbVD[i][3] ? dbVD[i][3].toString() : "",
          items: items,
          subtotal: parseFloat(dbVD[i][6]) || 0,
          iva: parseFloat(dbVD[i][7]) || 0,
          total: parseFloat(dbVD[i][8]) || 0,
          formaPago: dbVD[i][9] ? dbVD[i][9].toString() : "",
          estado: dbVD[i][10] ? dbVD[i][10].toString() : "Cobrada",
          notas: dbVD[i][11] ? dbVD[i][11].toString() : ""
        });
      }
    }

    // Pedidos a Producción
    const hojaPP = ss.getSheetByName("Pedidos_Produccion");
    if (hojaPP) {
      const dbPP = hojaPP.getDataRange().getValues();
      for (let i = 1; i < dbPP.length; i++) {
        if (!dbPP[i][0]) continue;
        let items = [];
        try { items = JSON.parse(dbPP[i][5] || "[]"); } catch(e) {}
        resultado.push({
          id: dbPP[i][0].toString(),
          fecha: dbPP[i][1] ? dbPP[i][1].toString() : "",
          tipo: "Pedido a Producción",
          clienteNombre: dbPP[i][2] ? dbPP[i][2].toString() : "",
          clienteDuiNit: dbPP[i][3] ? dbPP[i][3].toString() : "",
          items: items,
          subtotal: parseFloat(dbPP[i][6]) || 0,
          iva: parseFloat(dbPP[i][7]) || 0,
          total: parseFloat(dbPP[i][8]) || 0,
          formaPago: "Pendiente",
          estado: dbPP[i][9] ? dbPP[i][9].toString() : "Pendiente de Producción",
          notas: dbPP[i][10] ? dbPP[i][10].toString() : ""
        });
      }
    }

    // Cotizaciones convertidas
    const hojaCot = ss.getSheetByName("Cotizaciones");
    if (hojaCot) {
      const dbCot = hojaCot.getDataRange().getValues();
      for (let i = 1; i < dbCot.length; i++) {
        if (!dbCot[i][0] || (dbCot[i][10] || "").toString() !== "Convertida") continue;
        let items = [];
        try { items = JSON.parse(dbCot[i][6] || "[]"); } catch(e) {}
        resultado.push({
          id: dbCot[i][0].toString(),
          fecha: dbCot[i][1] ? dbCot[i][1].toString() : "",
          tipo: "Cotización Convertida",
          clienteNombre: dbCot[i][2] ? dbCot[i][2].toString() : "",
          clienteDuiNit: dbCot[i][3] ? dbCot[i][3].toString() : "",
          items: items,
          subtotal: parseFloat(dbCot[i][7]) || 0,
          iva: parseFloat(dbCot[i][8]) || 0,
          total: parseFloat(dbCot[i][9]) || 0,
          formaPago: "—",
          estado: "Convertida",
          notas: dbCot[i][11] ? dbCot[i][11].toString() : ""
        });
      }
    }

    resultado.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    return resultado;
  } catch (e) {
    console.error("Error al obtener historial de ventas: " + e.message);
    return [];
  }
}

// ==========================================================================
//   CANCELACIÓN Y DEVOLUCIÓN DE VENTAS
// ==========================================================================

/**
 * Cancela una venta directa o pedido a produccion:
 * 1. Marca la venta como "Cancelada" en su hoja
 * 2. Si tenia items con control de stock, genera WH-IN de devolucion
 *    en Recepciones y Movimientos_Inventario (regresa al inventario)
 * 3. Marca el ingreso correspondiente como "Cancelado" en Libro_Ingresos
 */
function cancelarVentaServidor(idVenta, motivo) {
  return ejecutarConLock(function() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      motivo = motivo || "Cancelación de venta";

      // Determinar si es VD o PED
      const esVD = idVenta.startsWith("VD-");
      const nombreHoja = esVD ? "Ventas_Directas" : "Pedidos_Produccion";
      const hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) return { exito: false, mensaje: "No se encontró la hoja de ventas." };

      const db = hoja.getDataRange().getValues();
      for (let i = 1; i < db.length; i++) {
        if (db[i][0].toString() !== idVenta) continue;

        const estadoActual = (db[i][esVD ? 10 : 9] || "").toString();
        if (estadoActual === "Cancelada") {
          return { exito: false, mensaje: "Esta venta ya fue cancelada." };
        }

        // 1. Marcar como Cancelada
        const colEstado = esVD ? 11 : 10;
        hoja.getRange(i + 1, colEstado).setValue("Cancelada");

        // 2. Revertir inventario si había items con stock
        let items = [];
        try { items = JSON.parse(db[i][5] || "[]"); } catch(e) {}

        if (esVD && items.length > 0) {
          const mapaVariantes = obtenerMapaVariantesConFlujoServidor();
          const mapaDesc = {};
          Object.keys(mapaVariantes).forEach(sku => {
            const d = (mapaVariantes[sku].descripcion || "").toLowerCase().trim();
            if (d) mapaDesc[d] = sku;
          });

          const hojaMov = obtenerOCrearHojaMovimientos(ss);
          const itemsConStock = [];

          items.forEach(item => {
            let sku = (item.sku || "").toString().trim();
            if (!sku || !mapaVariantes[sku]) {
              const d = (item.desc || item.descripcion || "").toLowerCase().trim();
              sku = mapaDesc[d] || "";
            }
            if (!sku) return;
            const info = mapaVariantes[sku];
            if (!info || !info.controlaStock) return;

            // Registrar entrada de devolucion en Movimientos_Inventario
            const idMov = "MOV-DEV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
            hojaMov.appendRow([
              idMov, new Date(), sku, "Entrada", "Devolución de venta",
              parseFloat(item.cant || item.cantidadRecibida || 0),
              "Sin Bodega",   // vuelve sin bodega especifica, el almacen la asigna
              idVenta, "Devolución: " + idVenta + " — " + motivo,
              Session.getActiveUser().getEmail() || ""
            ]);
            itemsConStock.push({
              sku, descripcion: item.desc || item.descripcion || sku,
              cantidadRecibida: parseFloat(item.cant || item.cantidadRecibida || 0),
              costoUnitario: parseFloat(item.prec || 0)
            });
          });

          // 3. Generar recibo WH-IN de devolución en Recepciones
          if (itemsConStock.length > 0) {
            let hojaRec = ss.getSheetByName("Recepciones");
            if (!hojaRec) {
              hojaRec = ss.insertSheet("Recepciones");
              hojaRec.appendRow(["ID", "Fecha", "ID_Orden", "Items (JSON)", "Usuario",
                "Tipo Destino", "Recibido Por", "Entrega Directa A", "Firma Digital",
                "Tipo Movimiento", "Bodega"]);
            }
            const idDevolucion = generarCorrelativoRecepcionServidor(hojaRec, "WH-IN");
            hojaRec.appendRow([
              idDevolucion, new Date(), idVenta,
              JSON.stringify(itemsConStock),
              resolverNombreUsuarioPorCorreoServidor(Session.getActiveUser().getEmail()),
              "Bodega", "", "Devolución de venta: " + idVenta, "",
              "Entrada", "Sin Bodega"
            ]);
          }
        }

        // 4. Cancelar en Libro_Ingresos si existe
        const hojaIng = ss.getSheetByName("Libro_Ingresos");
        if (hojaIng) {
          const dbIng = hojaIng.getDataRange().getValues();
          for (let j = 1; j < dbIng.length; j++) {
            if (dbIng[j][2].toString() === idVenta) {
              hojaIng.getRange(j + 1, 8).setValue("Cancelado");
              break;
            }
          }
        }

        registrarAuditoria("Ventas", "Cancelar", "Venta cancelada: " + idVenta + " — " + motivo);
        return {
          exito: true,
          mensaje: "¡Venta " + idVenta + " cancelada!" + (esVD && items.length > 0 ? " El inventario fue revertido." : ""),
          revirtioBodega: esVD
        };
      }
      return { exito: false, mensaje: "No se encontró la venta." };
    } catch(e) {
      return { exito: false, mensaje: "Error al cancelar: " + e.message };
    }
  });
}
