function extraerDatosBCP() {
  const REMITENTE = "notificaciones@notificacionesbcp.com.pe";
  const ETIQUETA = "tcprocesadabcp";
  const HOJA_NOMBRE = "Gastos Tarjeta BCP";
  const ENCABEZADOS = ["Fecha envío", "Monto (S/)", "Operación realizada", "Destinatario", "ID Mensaje"];

  const ASUNTOS_EXCLUIDOS = [
    "Constancia de Configuración de Tarjeta en Banca Móvil BCP - Servicio de Notificaciones BCP",
    "Se rechazó tu compra por e-commerce no permitido - Servicio de Notificaciones BCP",
    "Constancia de Transferencia Entre mis Cuentas - Servicio de Notificaciones BCP",
    "Constancia de transferencia propias",
    "Configuración de Tarjeta",
    "Se rechazó tu compra por E-Commerce no permitido - Servicio de Notificaciones BCP",
    "Realizamos una devolución de una operación a tu Tarjeta de Débito BCP - Servicio de Notificaciones BCP"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA_NOMBRE);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_NOMBRE);
    hoja.appendRow(ENCABEZADOS);
  }
  if (hoja.getLastRow() === 0) hoja.appendRow(ENCABEZADOS);

  let labelObj = GmailApp.getUserLabelByName(ETIQUETA);
  if (!labelObj) labelObj = GmailApp.createLabel(ETIQUETA);

  const query = `from:${REMITENTE} -label:${ETIQUETA}`;
  const threads = GmailApp.search(query);
  const messages = threads.flatMap(t => t.getMessages());

  const datosParaAnadir = [];

  const regexOperacion = /Operación realizada[\s\t:]*\*?(.+?)\*/i;

  const regexEmpresa   = /^Empresa[\s\t:]*([^\r\n]+)/im;
  const regexDestino   = /^Destino[\s\t:]*([^\r\n]+)/im;
  const regexPagadoA   = /^Pagado a[\s\t:]*([^\r\n]+)/im;
  const regexEnviadoA  = /^Enviado a[\s\t:]*([^\r\n]+)/im;

  for (const message of messages) {
    try {
      const thread = message.getThread();
      const idMensaje = message.getId();

      const asunto = message.getSubject() || "";

      if (ASUNTOS_EXCLUIDOS.includes(asunto.trim())) continue;

      if (thread.getLabels().some(l => l.getName() === ETIQUETA)) continue;

      const cuerpo = message.getPlainBody() || "";

      const fechaObj = message.getDate();
      const fechaTexto = fechaObj
        ? Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : "ERROR";

      // --- OPERACIÓN REALIZADA ---
      let operacion = "ERROR";
      const matchOperacion = cuerpo.match(regexOperacion);
      if (matchOperacion && matchOperacion[1]) {
        operacion = matchOperacion[1].trim();
        operacion = operacion.replace(/^\*+|\*+$/g, "").trim();
      }

      const opNorm = operacion.toLowerCase().trim();


      // ==========================================================
      //                  🔥 BLOQUE DE MONTOS ACTUALIZADO
      // ==========================================================

      let monto = "ERROR";

      // 1) Yapear / Yapeo
      if (opNorm === "yapear a celular" || opNorm === "yapeo a celular") {
        const m = cuerpo.match(/Monto enviado[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // 2) Pago de servicios
      else if (opNorm === "pago de servicios") {
        const m = cuerpo.match(/Monto total[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // 3) Pago de tarjeta propia BCP
      else if (opNorm === "pago de tarjeta propia bcp") {
        const m = cuerpo.match(/Monto pagado[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // 4) Aporte voluntario
      else if (opNorm === "aporte voluntario") {
        const m = cuerpo.match(/Total aportado[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // 5) Retiro
      else if (opNorm === "retiro") {
        const m = cuerpo.match(/(Monto retirado|Total retirado)[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[3].trim();
      }

      // 6) Transferencia a terceros BCP
      else if (opNorm === "transferencia a terceros bcp") {
        const m = cuerpo.match(/Monto transferido[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // 7) 🔥 CASO GENERAL (consumos en S/ o en $)
      else {
        const m = cuerpo.match(/Total\s+del\s+consumo[\s\S]*?(S\/|\$)\s*([\d.,]+)/i);
        if (m) monto = m[2].trim();
      }

      // ==========================================================


      // --- DESTINATARIO ---
      let destinatario = "ERROR";

      if (opNorm === "consumo tarjeta de crédito" || opNorm === "consumo tarjeta de débito") {
        const m = cuerpo.match(regexEmpresa);
        if (m && m[1]) destinatario = m[1].trim();
      }
      else if (opNorm === "aporte voluntario") {
        const m = cuerpo.match(regexDestino);
        if (m && m[1]) destinatario = m[1].trim();
      }
      else if (opNorm === "pago de tarjeta propia bcp" || opNorm === "yapear a celular") {
        const m = cuerpo.match(regexPagadoA);
        if (m && m[1]) destinatario = m[1].trim();
      }
      else if (opNorm === "yapeo a celular") {
        const m = cuerpo.match(regexEnviadoA);
        if (m && m[1]) destinatario = m[1].trim();
      }
      else if (opNorm === "pago de servicios") {
        const m = cuerpo.match(regexEmpresa);
        if (m && m[1]) destinatario = m[1].trim();
      }

      destinatario = destinatario.replace(/^\*+|\*+$/g, "").trim();

      if (destinatario === "ERROR") {
        const f1 = cuerpo.match(regexEmpresa);
        const f2 = cuerpo.match(regexDestino);
        const f3 = cuerpo.match(regexPagadoA);
        const f4 = cuerpo.match(regexEnviadoA);

        const fallback =
          (f1 && f1[1]) ||
          (f2 && f2[1]) ||
          (f3 && f3[1]) ||
          (f4 && f4[1]);

        if (fallback) destinatario = fallback.trim().replace(/^\*+|\*+$/g, "");
      }

      datosParaAnadir.push([fechaTexto, monto, operacion, destinatario, idMensaje]);
      thread.addLabel(labelObj);

    } catch (e) {
      datosParaAnadir.push([
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        "ERROR", "ERROR", "ERROR", message.getId() || "ERROR"
      ]);
    }
  }

  if (datosParaAnadir.length > 0) {
    hoja
      .getRange(hoja.getLastRow() + 1, 1, datosParaAnadir.length, ENCABEZADOS.length)
      .setValues(datosParaAnadir);
  }
}
