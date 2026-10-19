(function () {
  'use strict';

  // ---------------------------------------------------------------------
  // CONFIGURACIÓN
  //   API_URL: URL de la implementación de Apps Script (termina en /exec).
  //            Si está vacía, la app funciona en MODO DEMO (no guarda nada).
  //   TOKEN:   el mismo valor que pusiste en Code.gs.
  // ---------------------------------------------------------------------
  var CONFIG = {
    API_URL: '',
    TOKEN: 'a63f009e24790fc48f6d1f95',
    TIMEOUT_MS: 15000,               // lecturas (catálogo, nodos)
    TIMEOUT_POST_MS: 30000,          // envío de un registro (Apps Script puede tardar en guardarlo)
    BLOQUEAR_SI_NO_COINCIDE: false,  // true = no deja enviar si las sacas recolectadas no coinciden con las recibidas
    COMPARAR_CONTRA: 'ACUM'          // con qué volumen se compara lo recolectado: 'ACUM' (acumulado de la ventana, hoy 4 días) o 'D1' (ayer)
  };

  var MAX_SACAS = 1000;

  // ---------------------------------------------------------------------
  // PAÍSES (código del Excel -> nombre, idioma y zonas horarias del dispositivo)
  // ---------------------------------------------------------------------
  var PAISES = [
    { code: 'MLB', nombre: 'Brasil',    lang: 'pt-BR', region: 'BR' },
    { code: 'MLM', nombre: 'México',    lang: 'es-MX', region: 'MX' },
    { code: 'MLA', nombre: 'Argentina', lang: 'es-AR', region: 'AR' },
    { code: 'MLC', nombre: 'Chile',     lang: 'es-CL', region: 'CL' },
    { code: 'MCO', nombre: 'Colombia',  lang: 'es-CO', region: 'CO' }
  ];

  var ZONAS_HORARIAS = {
    MLB: ['America/Sao_Paulo', 'America/Manaus', 'America/Fortaleza', 'America/Bahia', 'America/Recife',
          'America/Belem', 'America/Cuiaba', 'America/Campo_Grande', 'America/Noronha', 'America/Rio_Branco',
          'America/Porto_Velho', 'America/Boa_Vista', 'America/Araguaina', 'America/Maceio',
          'America/Santarem', 'America/Eirunepe'],
    MLM: ['America/Mexico_City', 'America/Cancun', 'America/Merida', 'America/Monterrey', 'America/Matamoros',
          'America/Chihuahua', 'America/Ciudad_Juarez', 'America/Ojinaga', 'America/Hermosillo',
          'America/Mazatlan', 'America/Bahia_Banderas', 'America/Tijuana'],
    MLA: ['America/Buenos_Aires', 'America/Cordoba', 'America/Mendoza', 'America/Jujuy', 'America/Catamarca',
          'America/Rosario'],              // más todo lo que empiece con America/Argentina/
    MLC: ['America/Santiago', 'America/Punta_Arenas', 'Pacific/Easter'],
    MCO: ['America/Bogota']
  };

  // Sites por país (tomados de ARCANA.xlsx). Se usan si la hoja "Sites" está vacía o sin conexión.
  var SITES_BASE = {
    MLB: 'SAL1,SAM1,SBA1,SBA2,SBA3,SBA6,SCE1,SCE3,SC_ZS,SDF1,SDF2,SES1,SES2,SES3,SGO1,SGO2,SGO3,SJP1,SMG1,SMG2,SMG3,SMG4,SMG5,SMG6,SMG7,SMG8,SMG9,SMG10,SMG11,SMG12,SMG13,SMG14,SMG15,SMN1,SMR1,SMR2,SMS1,SMS2,SMS3,SPA1,SPE1,SPI1,SPR1,SPR2,SPR3,SPR4,SPR5,SPR6,SPR7,SPR8,SPR10,SPR11,SRD1,SRJ1,SRJ2,SRJ3,SRJ4,SRJ5,SRJ6,SRJ7,SRJ8,SRJ9,SRJ10,SRJ12,SRJ13,SRJ17,SRN1,SRS1,SRS2,SRS3,SRS4,SRS5,SRS7,SRS8,SRS9,SRS10,SSC1,SSC2,SSC3,SSC4,SSC5,SSC7,SSC8,SSC9,SSE1,SSP3,SSP4,SSP5,SSP6,SSP7,SSP8,SSP9,SSP10,SSP11,SSP12,SSP13,SSP14,SSP15,SSP16,SSP17,SSP18,SSP20,SSP21,SSP22,SSP23,SSP24,SSP25,SSP26,SSP27,SSP28,SSP29,SSP30,SSP31,SSP34,SSP36,SSP37,SSP38,SSP39,SSP40,SSP45,SSP46,SSP48,SSP49,SSP50,SSP51,SSP52,SSP53,SSP54,SSP55,SSP57,SSP64,STO1',
    MLM: 'SAG1,SBJ1,SCC1,SCD1,SCG1,SCH1,SCJ1,SCM1,SCN1,SCP1,SCQ1,SCT1,SCU1,SCV1,SCX1,SCY1,SDC1,SDG1,SEN1,SGD1,SGD2,SGD3,SHM1,SHP1,SJA1,SJD1,SLE1,SLP1,SLV1,SLW1,SLZ1,SMA1,SMD1,SMD3,SML1,SMO1,SMT1,SMT2,SMT3,SMX1,SMX2,SMX3,SMX4,SMX5,SMX7,SMX8,SMX9,SMX10,SMX11,SMX22,SMZ1,SNG1,SOX1,SPB1,SPB2,SPD1,SPV1,SPY1,SPZ1,SQR1,SQR2,SRX1,SSL1,STA1,STG1,STJ1,STL1,STN1,STP1,STR1,STT1,STX1,SUR1,SVH1,SVM1,SVR1,SXL1,SZC1,SZL1,SZM1',
    MLA: 'SBC1,SBU1,SBU2,SBU4,SBU5,SBU10,SCF2,SCF3,SCZ1,SLA1,SME1,SMQ1,SNQ1,SRE1,SRF1,SRO1,SRO2,STU1',
    MLC: 'SAF1,SBB1,SBB2,SLS1,SNU1,SPM1,SPO1,SRC1,SRM1,SRM2,SRM3,STC1,STM1,SVL1,SVP3',
    MCO: 'SAN1,SBL1,SBO1,SBY1,SCI1,SRI1,STI1,SVA1'
  };


  // ---------------------------------------------------------------------
  // IDIOMAS: español (base), portugués de Brasil, y pequeños ajustes regionales
  // ---------------------------------------------------------------------
  var I18N = {
    es: {
      doc_titulo: 'Recolección de sacas por nodo', titulo: 'Sacas por nodo', subtitulo: 'Registro de recolección',
      demo: 'Modo demo', offline: 'Sin conexión',
      footer_demo: 'Modo demo: los datos no se guardan', footer_live: 'Los registros se guardan en Google Sheets',
      pend_1: '1 registro espera conexión para enviarse.', pend_n: '{n} registros esperan conexión para enviarse.',
      reintentar: 'Reintentar',
      pais: 'País', pais_auto: 'Detectado automáticamente. Toca otro país para cambiarlo.',
      site: 'SVC', site_sel: 'Selecciona tu SVC', site_titulo: 'Selecciona tu SVC', site_buscar: 'Buscar SVC',
      site_cuenta: '{n} SVC', site_sin: 'Sin resultados', cerrar: 'Cerrar',
      nodo: 'Nodo', nodo_sel: 'Selecciona el nodo', nodo_primero: 'Primero elige el SVC', nodo_cargando: 'Cargando nodos…',
      nodo_reintentar: 'No se pudo cargar. Toca para reintentar', nodo_titulo: 'Selecciona el nodo', nodo_buscar: 'Buscar nodo',
      nodo_cuenta: '{n} nodos', nodo_usar: 'Usar “{q}”', nodo_hoy: 'Hoy', lista_extra: 'Ayer {a} · {d} días {b}',
      nodos_vacio: 'Este SVC no tiene nodos con sacas en los últimos {d} días. Puedes escribir el nodo.',
      e_nodos: 'No se pudo cargar la lista de nodos. Toca el campo para reintentar.',
      nodo_sin_datos: 'No hay volúmenes de este nodo. Registra lo que recolectes.',
      nodo_recibio: 'Sacas que recibió este nodo', ayer: 'Ayer', dias_acum: 'Últimos {d} días',
      corte: 'Volúmenes al {f}', corte_viejo: 'Volúmenes desactualizados: al {f}',
      sacas: 'Sacas recolectadas', sacas_hint: 'Solo números, de 1 a {max}.', menos: 'Restar una saca', mas: 'Sumar una saca',
      enviar: 'Enviar registro',
      ref_d1: 'ayer', ref_acum: 'en los últimos {d} días',
      cmp_ok: 'Coincide con las sacas recibidas {r}.', cmp_menos: 'Faltan {n} respecto a las {m} recibidas {r}.', cmp_mas: 'Hay {n} de más respecto a las {m} recibidas {r}.',
      e_no_coincide: 'La cantidad no coincide con las {m} sacas recibidas {r}. Revísala.',
      e_pais: 'Selecciona un país.', e_site: 'Selecciona tu SVC.', e_nodo: 'Selecciona el nodo.',
      e_sacas: 'Escribe la cantidad de sacas (solo números).', e_rango: 'La cantidad debe ser de 1 a {max}.',
      e_servidor: 'El servidor no aceptó el registro. Recarga la página y vuelve a intentarlo.',
      ok_titulo: 'Registro enviado', pend_titulo: 'Registro guardado en este teléfono',
      pend_nota: 'No hay conexión. Se enviará solo cuando vuelva el internet.',
      pend_nota_contacto: 'No se pudo contactar al servidor. Se guardó en este teléfono y se enviará solo.',
      pend_nota_tiempo: 'El servidor tardó demasiado en responder. Se guardó en este teléfono y se enviará solo.',
      pend_nota_servidor: 'El servidor respondió con un error. Se guardó en este teléfono y se enviará solo.',
      pend_codigo: 'Código: {c}', pend_ultimo: 'Último intento: {c}',
      r_site: 'SVC', r_nodo: 'Nodo', r_sacas: 'Recolectadas', r_ref: 'Recibió',
      ref_valor: '{a} ayer, {b} en {d} días', r_vs: 'Contra lo recibido',
      vs_ok: 'Coincide con las {m} recibidas {r}', vs_menos: '{n} menos que las {m} recibidas {r}', vs_mas: '{n} más que las {m} recibidas {r}',
      r_folio: 'Folio', nuevo: 'Nuevo registro',
      recientes: 'Últimos registros', hoy: 'Hoy: {n}',
      b_ok: 'Enviado', b_pend: 'Pendiente', b_err: 'Rechazado', saca_1: 'saca', saca_n: 'sacas'
    },
    pt: {
      doc_titulo: 'Coleta de sacas por nodo', titulo: 'Sacas por nodo', subtitulo: 'Registro de coleta',
      demo: 'Modo demo', offline: 'Sem conexão',
      footer_demo: 'Modo demo: os dados não são salvos', footer_live: 'Os registros são salvos no Google Sheets',
      pend_1: '1 registro aguarda conexão para ser enviado.', pend_n: '{n} registros aguardam conexão para serem enviados.',
      reintentar: 'Tentar de novo',
      pais: 'País', pais_auto: 'Detectado automaticamente. Toque em outro país para mudar.',
      site: 'SVC', site_sel: 'Selecione seu SVC', site_titulo: 'Selecione seu SVC', site_buscar: 'Buscar SVC',
      site_cuenta: '{n} SVC', site_sin: 'Nenhum resultado', cerrar: 'Fechar',
      nodo: 'Nodo', nodo_sel: 'Selecione o nodo', nodo_primero: 'Escolha o SVC primeiro', nodo_cargando: 'Carregando nodos…',
      nodo_reintentar: 'Não foi possível carregar. Toque para tentar de novo', nodo_titulo: 'Selecione o nodo', nodo_buscar: 'Buscar nodo',
      nodo_cuenta: '{n} nodos', nodo_usar: 'Usar “{q}”', nodo_hoy: 'Hoje', lista_extra: 'Ontem {a} · {d} dias {b}',
      nodos_vacio: 'Este SVC não tem nodos com sacas nos últimos {d} dias. Você pode digitar o nodo.',
      e_nodos: 'Não foi possível carregar a lista de nodos. Toque no campo para tentar de novo.',
      nodo_sin_datos: 'Não há volumes deste nodo. Registre o que coletar.',
      nodo_recibio: 'Sacas que este nodo recebeu', ayer: 'Ontem', dias_acum: 'Últimos {d} dias',
      corte: 'Volumes em {f}', corte_viejo: 'Volumes desatualizados: em {f}',
      sacas: 'Sacas coletadas', sacas_hint: 'Somente números, de 1 a {max}.', menos: 'Tirar uma saca', mas: 'Somar uma saca',
      enviar: 'Enviar registro',
      ref_d1: 'ontem', ref_acum: 'nos últimos {d} dias',
      cmp_ok: 'Confere com as sacas recebidas {r}.', cmp_menos: 'Faltam {n} em relação às {m} recebidas {r}.', cmp_mas: 'Há {n} a mais em relação às {m} recebidas {r}.',
      e_no_coincide: 'A quantidade não confere com as {m} sacas recebidas {r}. Revise.',
      e_pais: 'Selecione um país.', e_site: 'Selecione seu SVC.', e_nodo: 'Selecione o nodo.',
      e_sacas: 'Digite a quantidade de sacas (somente números).', e_rango: 'A quantidade deve ser de 1 a {max}.',
      e_servidor: 'O servidor não aceitou o registro. Recarregue a página e tente novamente.',
      ok_titulo: 'Registro enviado', pend_titulo: 'Registro salvo neste celular',
      pend_nota: 'Sem conexão. Será enviado automaticamente quando a internet voltar.',
      pend_nota_contacto: 'Não foi possível contatar o servidor. Salvo neste celular; será enviado automaticamente.',
      pend_nota_tiempo: 'O servidor demorou demais para responder. Salvo neste celular; será enviado automaticamente.',
      pend_nota_servidor: 'O servidor respondeu com erro. Salvo neste celular; será enviado automaticamente.',
      pend_codigo: 'Código: {c}', pend_ultimo: 'Última tentativa: {c}',
      r_site: 'SVC', r_nodo: 'Nodo', r_sacas: 'Coletadas', r_ref: 'Recebeu',
      ref_valor: '{a} ontem, {b} em {d} dias', r_vs: 'Contra o recebido',
      vs_ok: 'Confere com as {m} recebidas {r}', vs_menos: '{n} a menos das {m} recebidas {r}', vs_mas: '{n} a mais das {m} recebidas {r}',
      r_folio: 'Protocolo', nuevo: 'Novo registro',
      recientes: 'Últimos registros', hoy: 'Hoje: {n}',
      b_ok: 'Enviado', b_pend: 'Pendente', b_err: 'Rejeitado', saca_1: 'saca', saca_n: 'sacas'
    },
    // Ajustes regionales (sobre el español base)
    'es-AR': {
      pais_auto: 'Detectado automáticamente. Tocá otro país para cambiarlo.',
      site_sel: 'Seleccioná tu SVC', site_titulo: 'Seleccioná tu SVC', site_buscar: 'Buscá un SVC',
      nodo_sel: 'Seleccioná el nodo', nodo_primero: 'Primero elegí el SVC', nodo_titulo: 'Seleccioná el nodo', nodo_buscar: 'Buscá un nodo',
      nodo_reintentar: 'No se pudo cargar. Tocá para reintentar',
      nodos_vacio: 'Este SVC no tiene nodos con sacas en los últimos {d} días. Podés escribir el nodo.',
      e_nodos: 'No se pudo cargar la lista de nodos. Tocá el campo para reintentar.',
      nodo_sin_datos: 'No hay volúmenes de este nodo. Registrá lo que recolectes.',
      e_pais: 'Seleccioná un país.', e_site: 'Seleccioná tu SVC.', e_nodo: 'Seleccioná el nodo.',
      e_sacas: 'Escribí la cantidad de sacas (solo números).',
      e_no_coincide: 'La cantidad no coincide con las {m} sacas recibidas {r}. Revisala.',
      e_servidor: 'El servidor no aceptó el registro. Recargá la página y volvé a intentarlo.',
      pend_titulo: 'Registro guardado en este celular'
    },
    'es-MX': { pend_titulo: 'Registro guardado en este celular' },
    'es-CO': { pend_titulo: 'Registro guardado en este celular' }
  };

  var dict = I18N.es;

  function armarDiccionario(lang) {
    var base = lang === 'pt-BR' ? I18N.pt : I18N.es;
    var extra = I18N[lang] || {};
    var d = {}, k;
    for (k in base) d[k] = base[k];
    for (k in extra) d[k] = extra[k];
    return d;
  }

  function t(clave, vars) {
    var s = dict[clave];
    if (s === undefined) s = clave;
    var v = vars || {};
    return s.replace(/\{(\w+)\}/g, function (m, n) {
      if (v[n] !== undefined) return String(v[n]);
      if (n === 'max') return String(MAX_SACAS);
      if (n === 'd') return String(nodosDias);
      return m;
    });
  }

  // ---------------------------------------------------------------------
  // Utilidades
  // ---------------------------------------------------------------------
  var LS = { pais: 'reg.pais', site: 'reg.site', cola: 'reg.cola4', hist: 'reg.hist4', cat: 'reg.catalogo6', nodos: 'reg.nodos.' };
  var DEMO = !CONFIG.API_URL;
  // Sin API_URL la app NO funciona (para que nadie crea que está guardando). Para probar el diseño: agrega ?demo=1 al enlace.
  var DEMO_PERMITIDO = /[?&]demo=1(&|$)/.test(window.location.search);

  function $(id) { return document.getElementById(id); }
  function qsa(sel) { return Array.prototype.slice.call(document.querySelectorAll(sel)); }

  function leer(clave, defecto) {
    try { var v = localStorage.getItem(clave); return v ? JSON.parse(v) : defecto; }
    catch (e) { return defecto; }
  }
  function guardar(clave, valor) {
    try { localStorage.setItem(clave, JSON.stringify(valor)); } catch (e) { /* almacenamiento lleno o bloqueado */ }
  }
  function atraso(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function gerarId() {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') return window.crypto.randomUUID();
    var b = new Uint8Array(16);
    window.crypto.getRandomValues(b);
    return Array.prototype.map.call(b, function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
  }

  function limpiarTexto(s) {
    return String(s).replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function hora(ts) {
    var d = new Date(ts);
    return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
  }
  function esHoy(ts) {
    var a = new Date(ts), b = new Date();
    return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }
  function hoyISO() {
    var d = new Date();
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }

  function paisPorCodigo(code) {
    for (var i = 0; i < PAISES.length; i++) if (PAISES[i].code === code) return PAISES[i];
    return null;
  }

  function ordenNatural(a, b) { return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }); }

  // Petición con tiempo límite. Rechaza ante fallo de red, HTTP != 200 o respuesta que no sea JSON.
  function pedir(url, opciones, ms) {
    opciones = opciones || {};
    var ctl = typeof AbortController === 'function' ? new AbortController() : null;
    var timer = ctl ? setTimeout(function () { ctl.abort(); }, ms || CONFIG.TIMEOUT_MS) : null;
    if (ctl) opciones.signal = ctl.signal;
    return fetch(url, opciones)
      .then(function (r) {
        if (!r.ok) throw new Error('http ' + r.status);
        return r.json();
      })
      .then(function (j) { if (timer) clearTimeout(timer); return j; },
            function (e) { if (timer) clearTimeout(timer); throw e; });
  }

  // ---------------------------------------------------------------------
  // Detección del país
  //   Orden: ?pais=MLB en el enlace > país elegido a mano antes > zona horaria > idioma del teléfono
  // ---------------------------------------------------------------------
  function paisPorZonaHoraria() {
    var tz = '';
    try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; } catch (e) { tz = ''; }
    if (!tz) return null;
    if (tz.indexOf('America/Argentina/') === 0) return 'MLA';
    for (var code in ZONAS_HORARIAS) {
      if (ZONAS_HORARIAS[code].indexOf(tz) !== -1) return code;
    }
    return null;
  }

  function paisPorIdiomaDelTelefono() {
    var langs = navigator.languages && navigator.languages.length ? navigator.languages : [navigator.language || ''];
    for (var i = 0; i < langs.length; i++) {
      var m = /[-_]([A-Za-z]{2})$/.exec(langs[i] || '');
      if (!m) continue;
      var region = m[1].toUpperCase();
      for (var j = 0; j < PAISES.length; j++) if (PAISES[j].region === region) return PAISES[j].code;
    }
    return null;
  }

  function paisDeLaUrl() {
    try {
      var p = new URLSearchParams(window.location.search).get('pais');
      if (p) { p = p.toUpperCase(); if (paisPorCodigo(p)) return p; }
    } catch (e) { /* navegador sin URLSearchParams */ }
    return null;
  }

  function idiomaPorDefecto() {
    var l = (navigator.language || '').toLowerCase();
    return l.indexOf('pt') === 0 ? 'pt-BR' : 'es';
  }

  // ---------------------------------------------------------------------
  // Acceso a datos (contrato: devuelven Promise)
  // ---------------------------------------------------------------------
  function sitesBase() {
    var out = {};
    PAISES.forEach(function (p) { out[p.code] = SITES_BASE[p.code] ? SITES_BASE[p.code].split(',') : []; });
    return out;
  }

  // El catálogo solo trae los SVC (solo orígenes con datos). Si no se puede cargar, se usa la lista incluida en la app.
  function cargarCatalogo() {
    if (DEMO) return atraso(400).then(function () { return { sites: sitesBase() }; });
    var url = CONFIG.API_URL + '?action=catalogo&token=' + encodeURIComponent(CONFIG.TOKEN);
    return pedir(url).then(function (r) {
      if (!r || r.ok !== true) throw new Error('catalogo');
      var cat = { sites: r.sites || null };
      guardar(LS.cat, cat);
      return cat;
    }).catch(function () {
      return leer(LS.cat, null) || { sites: null }; // sin conexión: último catálogo conocido, o la lista incluida
    });
  }

  // Nodos que recibieron sacas desde un SVC, con volúmenes de ayer (d1) y del acumulado de la ventana de días (t3)
  function demoNodos(svc) {
    var h = 0, i;
    for (i = 0; i < svc.length; i++) h = (h * 31 + svc.charCodeAt(i)) % 997;
    var nodos = [];
    for (i = 0; i < 7; i++) {
      var d1 = (h + i * 7) % 25;
      nodos.push({ id: 'DEMO-' + svc + '-' + (i + 1), tipo: 'NODE', d1: d1, t3: d1 + ((h + i * 11) % 40) + 1 });
    }
    nodos.sort(function (a, b) { return (b.d1 - a.d1) || (b.t3 - a.t3); });
    return { nodos: nodos, corte: hoyISO(), dias: 4 };
  }

  function obtenerNodos(svc) {
    if (DEMO) return atraso(400).then(function () { return demoNodos(svc); });
    var url = CONFIG.API_URL + '?action=nodos&svc=' + encodeURIComponent(svc) + '&token=' + encodeURIComponent(CONFIG.TOKEN);
    return pedir(url).then(function (r) {
      if (!r || r.ok !== true || !r.nodos) throw new Error('nodos');
      var out = { nodos: r.nodos, corte: r.corte || '', dias: Number(r.dias) || 4 };
      guardar(LS.nodos + svc, out);
      return out;
    }).catch(function (e) {
      var cache = leer(LS.nodos + svc, null);
      if (cache) return cache; // sin conexión: usa la última lista conocida de ese SVC
      throw e;
    });
  }

  // Explica por qué falló un envío: sin red, sin respuesta, tiempo agotado o error del servidor.
  function clasificarFallo(e) {
    var nombre = e && e.name, msg = (e && e.message) || '', m = /^http (\d+)/.exec(msg);
    if (nombre === 'AbortError') return { tipo: 'tiempo', cod: 'TIMEOUT' };
    if (m) return { tipo: 'servidor', cod: 'HTTP ' + m[1] };
    if (nombre === 'SyntaxError') return { tipo: 'servidor', cod: 'RESPUESTA NO VALIDA' };
    if (navigator.onLine === false) return { tipo: 'red', cod: 'SIN RED' };
    return { tipo: 'contacto', cod: 'SIN RESPUESTA' };
  }

  // Errores del servidor que no son culpa del dato: el registro debe quedar en cola y reintentarse
  function esTransitorio(res) {
    return !!res && res.ok !== true && (res.error === 'ocupado' || res.error === 'interno');
  }

  // Resuelve {ok:true} si se guardó, {ok:false,error} si el servidor lo rechazó.
  // Rechaza (catch) solo ante problemas de red: esos casos se encolan.
  function enviarRegistro(reg) {
    if (DEMO) return atraso(800).then(function () { return { ok: true }; });
    // Se envía por GET (action=registrar) y no por POST: Apps Script contesta un POST con una redirección que solo admite
    // GET y algunos navegadores o redes reintentan con POST, lo que da "HTTP 405" aunque el registro sí se guardó.
    // Un reintento es seguro: el servidor ignora un folio (id) que ya guardó.
    var qs = [
      'action=registrar',
      'token=' + encodeURIComponent(CONFIG.TOKEN),
      'id=' + encodeURIComponent(reg.id),
      'pais=' + encodeURIComponent(reg.pais),
      'site=' + encodeURIComponent(reg.site),
      'nodo=' + encodeURIComponent(reg.nodo),
      'sacas=' + encodeURIComponent(reg.sacas),
      'creado=' + encodeURIComponent(reg.creado)
    ].join('&');
    return pedir(CONFIG.API_URL + '?' + qs, { cache: 'no-store' }, CONFIG.TIMEOUT_POST_MS);
  }

  // ---------------------------------------------------------------------
  // Estado y referencias
  // ---------------------------------------------------------------------
  var el = {
    form: $('form'), paises: $('paises'), paisNota: $('pais-nota'), siteBtn: $('site-btn'),
    nodoBtn: $('nodo-btn'), sacas: $('sacas'),
    enviar: $('enviar'), menos: $('menos'), mas: $('mas'),
    erroGeral: $('erro-geral'), erroCatalogo: $('erro-catalogo'),
    paisErr: $('pais-error'), siteErr: $('site-error'),
    nodoErr: $('nodo-error'), sacasErr: $('sacas-error'), nodoNota: $('nodo-nota'),
    nodoInfo: $('nodo-info'), nodoD1: $('nodo-d1'), nodo3d: $('nodo-3d'), nodoLblAcum: $('nodo-lbl-acum'), nodoCorte: $('nodo-corte'), cmp: $('cmp'),
    formCard: $('form-card'), exitoCard: $('sucesso-card'),
    exitoIcono: $('exito-icono'), exitoTitulo: $('exito-titulo'), exitoNota: $('exito-nota'),
    rSite: $('resumo-site'), rNodo: $('resumo-nodo'), rSacas: $('resumo-sacas'),
    filaRef: $('fila-ref'), rRef: $('resumo-ref'), filaVs: $('fila-vs'), rVs: $('resumo-vs'), rId: $('resumo-id'),
    novo: $('novo'),
    estadoTag: $('estado-tag'), footer: $('footer-texto'),
    pendientes: $('pendientes'), pendientesTexto: $('pendientes-texto'), reintentar: $('reintentar'),
    recientes: $('recientes'), recientesLista: $('recientes-lista'), recientesHoy: $('recientes-hoy'),
    sheet: $('sheet'), sheetTitle: $('sheet-title'), sheetBackdrop: $('sheet-backdrop'), sheetClose: $('sheet-close'),
    siteBuscar: $('site-buscar'), siteCuenta: $('site-cuenta'), siteLista: $('site-lista'),
    segPais: $('seg-pais'), segSite: $('seg-site'), segNodo: $('seg-nodo'), segSacas: $('seg-sacas')
  };

  var sitesPorPais = sitesBase();
  Object.keys(sitesPorPais).forEach(function (k) { sitesPorPais[k].sort(ordenNatural); });

  var paisActual = '';
  var paisAuto = false;            // true si el país se detectó solo y la persona no lo ha cambiado
  var siteActual = '';             // SVC

  var nodosEstado = 'inactivo';    // inactivo | cargando | listo | vacio | error
  var nodosLista = [];
  var nodosDias = 4;               // ventana de días del acumulado (la reporta el servidor: SACAS_4D -> 4)
  var nodosCorte = '';
  var nodosPeticion = 0;           // evita que una respuesta lenta pise a una más nueva
  var nodoActual = null;           // { id, d1, t3 }  (d1 y t3 son null si el nodo se escribió a mano)

  var sheetModo = 'site';          // site | nodo
  var enviando = false;
  var vaciando = false;
  var ultimoFallo = '';            // código del último intento fallido de enviar la cola (para diagnóstico)
  var focoPrevio = null;

  // ---------------------------------------------------------------------
  // Idioma
  // ---------------------------------------------------------------------
  function pintarPaisNota() {
    el.paisNota.hidden = !(paisAuto && paisActual);
    el.paisNota.textContent = t('pais_auto');
  }

  function pintarSiteBtn() {
    el.siteBtn.textContent = siteActual || t('site_sel');
    el.siteBtn.classList.toggle('is-placeholder', !siteActual);
    el.siteBtn.disabled = !paisActual;
  }

  function pintarTituloSheet() {
    var nodo = sheetModo === 'nodo';
    el.sheetTitle.textContent = t(nodo ? 'nodo_titulo' : 'site_titulo');
    el.siteBuscar.setAttribute('placeholder', t(nodo ? 'nodo_buscar' : 'site_buscar'));
  }

  function aplicarIdioma(lang) {
    dict = armarDiccionario(lang);
    document.documentElement.lang = lang;
    document.title = t('doc_titulo');
    qsa('[data-i18n]').forEach(function (n) { n.textContent = t(n.getAttribute('data-i18n')); });
    qsa('[data-i18n-ph]').forEach(function (n) { n.setAttribute('placeholder', t(n.getAttribute('data-i18n-ph'))); });
    qsa('[data-i18n-aria]').forEach(function (n) { n.setAttribute('aria-label', t(n.getAttribute('data-i18n-aria'))); });

    pintarPaisNota();
    pintarSiteBtn();
    pintarNodoBtn();
    pintarNodoNota();
    pintarNodoInfo();
    pintarComparacion();
    pintarConexion();
    pintarPendientes();
    pintarHist();
    pintarTituloSheet();
    if (!el.sheet.hidden) pintarLista();
  }

  // ---------------------------------------------------------------------
  // País
  // ---------------------------------------------------------------------
  function pintarPaises() {
    var frag = document.createDocumentFragment();
    PAISES.forEach(function (p) {
      var label = document.createElement('label');
      label.className = 'chip';
      var input = document.createElement('input');
      input.type = 'radio';
      input.name = 'pais';
      input.value = p.code;
      var span = document.createElement('span');
      var nombre = document.createElement('b');
      nombre.textContent = p.nombre;
      var codigo = document.createElement('small');
      codigo.textContent = p.code;
      span.appendChild(nombre);
      span.appendChild(codigo);
      label.appendChild(input);
      label.appendChild(span);
      frag.appendChild(label);
    });
    el.paises.replaceChildren(frag);
  }

  function seleccionarPais(code, manual) {
    var p = paisPorCodigo(code);
    if (!p) return;
    var cambio = paisActual !== code;
    paisActual = code;
    qsa('#paises input').forEach(function (r) { r.checked = r.value === code; });

    if (manual) {
      paisAuto = false;
      guardar(LS.pais, code);
    }
    // Un SVC pertenece a un solo país: si cambia el país, hay que volver a elegirlo
    if (siteActual && (sitesPorPais[code] || []).indexOf(siteActual) === -1) {
      siteActual = '';
      reiniciarNodos();
    }
    if (cambio) limpiarErrores();
    aplicarIdioma(p.lang);
    actualizarAvance();
  }

  // ---------------------------------------------------------------------
  // Nodos del SVC elegido
  // ---------------------------------------------------------------------
  function reiniciarNodos() {
    nodosPeticion++;
    nodosEstado = 'inactivo';
    nodosLista = [];
    nodosCorte = '';
    nodoActual = null;
    pintarNodoBtn();
    pintarNodoNota();
    pintarNodoInfo();
    actualizarAvance();
  }

  function cargarNodos(svc) {
    var mia = ++nodosPeticion;
    nodosEstado = 'cargando';
    nodosLista = [];
    nodosCorte = '';
    nodoActual = null;
    pintarNodoBtn();
    pintarNodoNota();
    pintarNodoInfo();
    actualizarAvance();
    obtenerNodos(svc)
      .then(function (r) {
        if (mia !== nodosPeticion) return;
        nodosLista = r.nodos || [];
        nodosCorte = r.corte || '';
        nodosDias = Number(r.dias) || 4;
        nodosEstado = nodosLista.length ? 'listo' : 'vacio';
      })
      .catch(function () {
        if (mia !== nodosPeticion) return;
        nodosEstado = 'error';
      })
      .then(function () {
        if (mia !== nodosPeticion) return;
        pintarNodoBtn();
        pintarNodoNota();
        pintarNodoInfo();
        actualizarAvance();
      });
  }

  function pintarNodoBtn() {
    var texto, deshabilitado = false, marcador = true;
    if (!siteActual) { texto = t('nodo_primero'); deshabilitado = true; }
    else if (nodosEstado === 'cargando') { texto = t('nodo_cargando'); deshabilitado = true; }
    else if (nodosEstado === 'error') { texto = t('nodo_reintentar'); }
    else if (nodoActual) { texto = nodoActual.id; marcador = false; }
    else { texto = t('nodo_sel'); }
    el.nodoBtn.textContent = texto;
    el.nodoBtn.classList.toggle('is-placeholder', marcador);
    el.nodoBtn.disabled = deshabilitado;
  }

  function pintarNodoNota() {
    var msg = '';
    if (siteActual) {
      if (nodosEstado === 'error') msg = t('e_nodos');
      else if (nodosEstado === 'vacio' && !nodoActual) msg = t('nodos_vacio');
      else if (nodoActual && typeof nodoActual.d1 !== 'number') msg = t('nodo_sin_datos');
    }
    el.nodoNota.hidden = !msg;
    el.nodoNota.textContent = msg;
  }

  function fechaCorta(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
    return m ? m[3] + '/' + m[2] : iso;
  }

  // La extracción lleva su fecha de corte; si va más de un día atrás se avisa
  function corteViejo(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
    if (!m) return false;
    var corte = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    var hoy = new Date(); hoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    return Math.round((hoy - corte) / 86400000) >= 2;
  }

  function pintarNodoInfo() {
    var n = nodoActual;
    var tiene = !!n && typeof n.d1 === 'number';
    el.nodoInfo.hidden = !tiene;
    if (!tiene) return;
    el.nodoD1.textContent = String(n.d1);
    el.nodo3d.textContent = String(n.t3);
    el.nodoLblAcum.textContent = t('dias_acum');
    var viejo = nodosCorte && corteViejo(nodosCorte);
    el.nodoCorte.hidden = !nodosCorte;
    el.nodoCorte.classList.toggle('is-viejo', !!viejo);
    el.nodoCorte.textContent = nodosCorte ? t(viejo ? 'corte_viejo' : 'corte', { f: fechaCorta(nodosCorte) }) : '';
  }

  function elegirNodo(n) {
    nodoActual = n;
    el.nodoBtn.removeAttribute('aria-invalid');
    el.nodoErr.hidden = true;
    el.nodoErr.textContent = '';
    pintarNodoBtn();
    pintarNodoNota();
    pintarNodoInfo();
    cerrarSheet();
    actualizarAvance();
    el.sacas.focus();
  }

  function registradoHoy(id) {
    return leer(LS.hist, []).some(function (x) { return x.site === siteActual && x.nodo === id && esHoy(x.ts); });
  }

  // Volumen de referencia contra el que se compara lo recolectado
  function referencia() {
    var n = nodoActual;
    if (!n || typeof n.d1 !== 'number') return null;
    var d1 = CONFIG.COMPARAR_CONTRA === 'D1';
    return { valor: d1 ? n.d1 : n.t3, r: t(d1 ? 'ref_d1' : 'ref_acum') };
  }

  function textoComparacion(sacas, ref, claves) {
    var dif = sacas - ref.valor;
    return dif === 0 ? t(claves[0], { m: ref.valor, r: ref.r })
      : dif < 0 ? t(claves[1], { n: -dif, m: ref.valor, r: ref.r }) : t(claves[2], { n: dif, m: ref.valor, r: ref.r });
  }

  function pintarComparacion() {
    var ref = referencia();
    if (!ref || !sacasValidas()) { el.cmp.hidden = true; el.cmp.textContent = ''; return; }
    var dif = Number(el.sacas.value) - ref.valor;
    el.cmp.hidden = false;
    el.cmp.className = 'cmp ' + (dif === 0 ? 'cmp--ok' : 'cmp--warn');
    el.cmp.textContent = textoComparacion(Number(el.sacas.value), ref, ['cmp_ok', 'cmp_menos', 'cmp_mas']);
  }

  // ---------------------------------------------------------------------
  // Hoja de selección (sirve para SVC y para nodos)
  // ---------------------------------------------------------------------
  function pintarLista() {
    var nodo = sheetModo === 'nodo';
    var q = limpiarTexto(el.siteBuscar.value).toUpperCase();
    var base = nodo
      ? nodosLista.map(function (n) { return { valor: n.id, extra: t('lista_extra', { a: n.d1, b: n.t3 }), hoy: registradoHoy(n.id) }; })
      : (sitesPorPais[paisActual] || []).map(function (s) { return { valor: s }; });
    var filtrados = q ? base.filter(function (x) { return x.valor.toUpperCase().indexOf(q) !== -1; }) : base;
    var actual = nodo ? (nodoActual && nodoActual.id) : siteActual;

    el.siteCuenta.textContent = t(nodo ? 'nodo_cuenta' : 'site_cuenta', { n: filtrados.length });
    var frag = document.createDocumentFragment();

    filtrados.forEach(function (x) {
      var li = document.createElement('li');
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'sheet__item';
      b.setAttribute('data-valor', x.valor);
      if (x.valor === actual) b.setAttribute('aria-current', 'true');

      var main = document.createElement('span');
      main.className = 'sheet__main';
      var txt = document.createElement('span');
      txt.textContent = x.valor;
      main.appendChild(txt);
      if (x.hoy) {
        var bd = document.createElement('span');
        bd.className = 'badge--hoy';
        bd.textContent = '✓ ' + t('nodo_hoy');
        main.appendChild(bd);
      }
      b.appendChild(main);

      var der = document.createElement('span');
      der.className = 'sheet__extra';
      der.textContent = (x.extra || '') + (x.valor === actual ? ' ✓' : '');
      b.appendChild(der);

      li.appendChild(b);
      frag.appendChild(li);
    });

    // Si el nodo no está en la lista, se puede escribir (queda sin volúmenes de referencia)
    var manual = nodo ? q.replace(/[^A-Z0-9_.\-]/g, '') : '';
    var yaEsta = filtrados.some(function (x) { return x.valor.toUpperCase() === manual; });
    if (nodo && manual.length >= 3 && !yaEsta) {
      var lim = document.createElement('li');
      var bm = document.createElement('button');
      bm.type = 'button';
      bm.className = 'sheet__item sheet__manual';
      bm.setAttribute('data-manual', manual);
      var tm = document.createElement('span');
      tm.textContent = t('nodo_usar', { q: manual });
      bm.appendChild(tm);
      lim.appendChild(bm);
      frag.appendChild(lim);
    } else if (!filtrados.length) {
      var vacio = document.createElement('li');
      vacio.className = 'sheet__empty';
      vacio.textContent = t('site_sin');
      frag.appendChild(vacio);
    }
    el.siteLista.replaceChildren(frag);
  }

  function abrirSheet(modo) {
    if (modo === 'site' && !paisActual) return;
    if (modo === 'nodo' && !siteActual) return;
    sheetModo = modo;
    focoPrevio = document.activeElement;
    el.siteBuscar.value = '';
    pintarTituloSheet();
    pintarLista();
    el.sheet.hidden = false;
    document.body.style.overflow = 'hidden';
    el.siteBuscar.focus();
    var sel = el.siteLista.querySelector('[aria-current="true"]');
    if (sel && sel.scrollIntoView) sel.scrollIntoView({ block: 'center' });
  }

  function cerrarSheet() {
    el.sheet.hidden = true;
    document.body.style.overflow = '';
    if (focoPrevio && focoPrevio.focus) focoPrevio.focus();
  }

  function elegirSite(s) {
    var cambio = s !== siteActual;
    siteActual = s;
    el.siteBtn.removeAttribute('aria-invalid');
    el.siteErr.hidden = true;
    el.siteErr.textContent = '';
    pintarSiteBtn();
    cerrarSheet();
    if (cambio || nodosEstado === 'inactivo' || nodosEstado === 'error') cargarNodos(s);
    actualizarAvance();
  }

  function atraparTab(ev) {
    if (ev.key === 'Escape') { cerrarSheet(); return; }
    if (ev.key !== 'Tab') return;
    var f = qsa('#sheet-panel button, #sheet-panel input');
    if (!f.length) return;
    var primero = f[0], ultimo = f[f.length - 1];
    if (ev.shiftKey && document.activeElement === primero) { ev.preventDefault(); ultimo.focus(); }
    else if (!ev.shiftKey && document.activeElement === ultimo) { ev.preventDefault(); primero.focus(); }
  }

  // ---------------------------------------------------------------------
  // Catálogo y formulario
  // ---------------------------------------------------------------------
  function pintarCatalogo(cat) {
    // Si el servidor trae SVC, reemplazan a los del Excel
    if (cat.sites) {
      var total = 0;
      PAISES.forEach(function (p) { total += (cat.sites[p.code] || []).length; });
      if (total > 0) {
        sitesPorPais = {};
        PAISES.forEach(function (p) { sitesPorPais[p.code] = (cat.sites[p.code] || []).slice().sort(ordenNatural); });
        if (siteActual && (sitesPorPais[paisActual] || []).indexOf(siteActual) === -1) { siteActual = ''; reiniciarNodos(); }
      }
    }

    pintarSiteBtn();
  }

  function sacasValidas() {
    var n = el.sacas.value;
    return /^[0-9]{1,4}$/.test(n) && Number(n) >= 1 && Number(n) <= MAX_SACAS;
  }

  function actualizarAvance() {
    el.segPais.classList.toggle('is-on', !!paisActual);
    el.segSite.classList.toggle('is-on', !!siteActual);
    el.segNodo.classList.toggle('is-on', !!nodoActual);
    el.segSacas.classList.toggle('is-on', sacasValidas());
    pintarComparacion();
  }

  function marcarError(campo, pNode, msg) {
    pNode.textContent = msg;
    pNode.hidden = false;
    if (campo) campo.setAttribute('aria-invalid', 'true');
  }

  function limpiarCampo(campo, pNode) {
    if (campo) campo.removeAttribute('aria-invalid');
    pNode.hidden = true;
    pNode.textContent = '';
  }

  function limpiarErrores() {
    [el.paisErr, el.siteErr, el.nodoErr, el.sacasErr, el.erroGeral].forEach(function (n) {
      n.hidden = true; n.textContent = '';
    });
    [el.siteBtn, el.nodoBtn, el.sacas].forEach(function (c) { c.removeAttribute('aria-invalid'); });
    el.paises.classList.remove('is-invalid');
  }

  function validar() {
    var ok = true;
    var n = el.sacas.value;
    var primero = null;

    if (!paisPorCodigo(paisActual)) {
      marcarError(null, el.paisErr, t('e_pais'));
      el.paises.classList.add('is-invalid');
      primero = primero || el.paises.querySelector('input'); ok = false;
    }
    if (!siteActual || (sitesPorPais[paisActual] || []).indexOf(siteActual) === -1) {
      marcarError(el.siteBtn, el.siteErr, t('e_site'));
      primero = primero || el.siteBtn; ok = false;
    }
    if (!nodoActual) {
      marcarError(el.nodoBtn, el.nodoErr, t('e_nodo'));
      primero = primero || (el.nodoBtn.disabled ? null : el.nodoBtn); ok = false;
    }
    if (!/^[0-9]{1,4}$/.test(n)) {
      marcarError(el.sacas, el.sacasErr, t('e_sacas'));
      primero = primero || el.sacas; ok = false;
    } else if (Number(n) < 1 || Number(n) > MAX_SACAS) {
      marcarError(el.sacas, el.sacasErr, t('e_rango'));
      primero = primero || el.sacas; ok = false;
    } else {
      var ref = referencia();
      if (CONFIG.BLOQUEAR_SI_NO_COINCIDE && ref && Number(n) !== ref.valor) {
        marcarError(el.sacas, el.sacasErr, t('e_no_coincide', { m: ref.valor, r: ref.r }));
        primero = primero || el.sacas; ok = false;
      }
    }
    if (primero) primero.focus();
    return ok ? { pais: paisActual, site: siteActual, nodo: nodoActual, sacas: Number(n) } : null;
  }

  // Solo dígitos, sin ceros a la izquierda, máximo 4 cifras
  function soloDigitos() {
    var limpio = el.sacas.value.replace(/[^0-9]/g, '').replace(/^0+(?=[0-9])/, '').slice(0, 4);
    if (limpio !== el.sacas.value) el.sacas.value = limpio;
    actualizarAvance();
  }

  function sumarSacas(n) {
    var v = (parseInt(el.sacas.value, 10) || 0) + n;
    if (v > MAX_SACAS) v = MAX_SACAS;
    el.sacas.value = v < 1 ? '' : String(v);
    limpiarCampo(el.sacas, el.sacasErr);
    actualizarAvance();
  }

  function setCargando(on) {
    enviando = on;
    el.enviar.disabled = on;
    el.enviar.classList.toggle('is-loading', on);
  }

  // ---------------------------------------------------------------------
  // Cola sin conexión e historial
  // ---------------------------------------------------------------------
  function encolar(reg) {
    var cola = leer(LS.cola, []);
    cola.push(reg);
    guardar(LS.cola, cola);
  }

  function guardarHist(reg, estado) {
    var h = leer(LS.hist, []);
    h.unshift({ id: reg.id, pais: reg.pais, site: reg.site, nodo: reg.nodo, sacas: reg.sacas, ts: reg.creado, estado: estado });
    guardar(LS.hist, h.slice(0, 30));
    depurarHist();
  }

  function marcarHist(id, estado) {
    var h = leer(LS.hist, []);
    h.forEach(function (x) { if (x.id === id) x.estado = estado; });
    guardar(LS.hist, h);
  }

  function pintarPendientes() {
    var n = leer(LS.cola, []).length;
    el.pendientes.hidden = n === 0;
    if (n) el.pendientesTexto.textContent = (n === 1 ? t('pend_1') : t('pend_n', { n: n })) + (ultimoFallo ? ' (' + t('pend_ultimo', { c: ultimoFallo }) + ')' : '');
  }

  // Solo se guardan y se muestran los registros de HOY (la lista vive en el teléfono; no se descarga de internet).
  // Los de días anteriores se borran, salvo los que siguen en la cola esperando conexión para enviarse.
  function depurarHist() {
    var h = leer(LS.hist, []);
    var enCola = {};
    leer(LS.cola, []).forEach(function (r) { enCola[r.id] = true; });
    var vigentes = h.filter(function (x) { return esHoy(x.ts) || (x.estado === 'pendiente' && enCola[x.id]); });
    if (vigentes.length !== h.length) guardar(LS.hist, vigentes);
  }

  function pintarHist() {
    var h = leer(LS.hist, []).filter(function (x) { return esHoy(x.ts); });
    el.recientes.hidden = h.length === 0;
    if (!h.length) return;
    el.recientesHoy.textContent = t('hoy', { n: h.length });

    var frag = document.createDocumentFragment();
    h.slice(0, 5).forEach(function (x) {
      var li = document.createElement('li');
      li.className = 'recent__item';

      var main = document.createElement('div');
      main.className = 'recent__main';
      var nombre = document.createElement('div');
      nombre.className = 'recent__name';
      nombre.textContent = x.nodo;
      var meta = document.createElement('div');
      meta.className = 'recent__meta';
      meta.textContent = [x.site, x.sacas + ' ' + (x.sacas === 1 ? t('saca_1') : t('saca_n')), hora(x.ts)].join(', ');
      main.appendChild(nombre);
      main.appendChild(meta);

      var badge = document.createElement('span');
      var clase = x.estado === 'enviado' ? 'ok' : x.estado === 'pendiente' ? 'pend' : 'err';
      badge.className = 'badge badge--' + clase;
      badge.textContent = x.estado === 'enviado' ? t('b_ok') : x.estado === 'pendiente' ? t('b_pend') : t('b_err');

      li.appendChild(main);
      li.appendChild(badge);
      frag.appendChild(li);
    });
    el.recientesLista.replaceChildren(frag);
  }

  function vaciarCola() {
    if (vaciando || DEMO) return;
    if (!leer(LS.cola, []).length) { pintarPendientes(); return; }
    vaciando = true;
    (function siguiente() {
      var cola = leer(LS.cola, []);
      if (!cola.length) { vaciando = false; pintarPendientes(); pintarHist(); return; }
      var reg = cola[0];
      enviarRegistro(reg)
        .then(function (res) {
          if (esTransitorio(res)) {          // el servidor falló: se conserva en la cola y se reintenta después
            ultimoFallo = String(res.error).toUpperCase() + (res.detalle ? ': ' + String(res.detalle).slice(0, 60) : '');
            vaciando = false; pintarPendientes(); return;
          }
          ultimoFallo = '';
          marcarHist(reg.id, res && res.ok === true ? 'enviado' : 'rechazado');
          guardar(LS.cola, leer(LS.cola, []).slice(1));
          pintarPendientes();
          pintarHist();
          siguiente();
        })
        .catch(function (e) { ultimoFallo = clasificarFallo(e).cod; vaciando = false; pintarPendientes(); });
    })();
  }

  // ---------------------------------------------------------------------
  // Envío
  // ---------------------------------------------------------------------
  function mostrarExito(reg, estado, fallo) {
    var pendiente = estado === 'pendiente';
    var p = paisPorCodigo(reg.pais);
    el.exitoIcono.classList.toggle('is-pending', pendiente);
    el.exitoTitulo.textContent = pendiente ? t('pend_titulo') : t('ok_titulo');
    el.exitoNota.hidden = !pendiente;
    var nota = '';
    if (pendiente) {
      var tipo = fallo ? fallo.tipo : 'red';
      nota = t(tipo === 'tiempo' ? 'pend_nota_tiempo' : tipo === 'servidor' ? 'pend_nota_servidor' : tipo === 'contacto' ? 'pend_nota_contacto' : 'pend_nota');
      if (fallo && fallo.tipo !== 'red') nota += ' (' + t('pend_codigo', { c: fallo.cod }) + ')';
    }
    el.exitoNota.textContent = nota;
    el.rSite.textContent = reg.site + (p ? ' (' + p.nombre + ')' : '');
    el.rNodo.textContent = reg.nodo;
    el.rSacas.textContent = String(reg.sacas);
    var tiene = typeof reg.d1 === 'number';
    el.filaRef.hidden = !tiene;
    el.filaVs.hidden = !tiene;
    if (tiene) {
      el.rRef.textContent = t('ref_valor', { a: reg.d1, b: reg.t3 });
      var d1 = CONFIG.COMPARAR_CONTRA === 'D1';
      el.rVs.textContent = textoComparacion(reg.sacas, { valor: d1 ? reg.d1 : reg.t3, r: t(d1 ? 'ref_d1' : 'ref_acum') }, ['vs_ok', 'vs_menos', 'vs_mas']);
    }
    el.rId.textContent = reg.id.slice(0, 8).toUpperCase();
    el.formCard.hidden = true;
    el.exitoCard.hidden = false;
    window.scrollTo(0, 0);
  }

  function aoSubmit(ev) {
    ev.preventDefault();
    if (enviando) return;
    limpiarErrores();
    var d = validar();
    if (!d) return;

    var reg = {
      id: gerarId(),
      pais: d.pais, site: d.site, nodo: d.nodo.id, sacas: d.sacas,
      d1: typeof d.nodo.d1 === 'number' ? d.nodo.d1 : null,   // solo para mostrar; el servidor los lee de la extracción
      t3: typeof d.nodo.t3 === 'number' ? d.nodo.t3 : null,
      creado: new Date().toISOString()
    };
    guardar(LS.site, { pais: reg.pais, site: reg.site });
    setCargando(true);

    enviarRegistro(reg)
      .then(function (res) {
        if (res && res.ok === true) {
          ultimoFallo = '';
          guardarHist(reg, 'enviado');
          mostrarExito(reg, 'enviado');
        } else if (esTransitorio(res)) {
          // El servidor falló por su cuenta: no se pierde el dato, se guarda y se reintenta
          ultimoFallo = String(res.error).toUpperCase() + (res.detalle ? ': ' + String(res.detalle).slice(0, 60) : '');
          encolar(reg);
          guardarHist(reg, 'pendiente');
          mostrarExito(reg, 'pendiente', { tipo: 'servidor', cod: ultimoFallo });
        } else {
          // El servidor rechazó el dato (campos, sacas, SVC, token...): se muestra el motivo
          el.erroGeral.textContent = t('e_servidor') + ' (' + String((res && res.error) || 'sin detalle') + ')';
          el.erroGeral.hidden = false;
        }
      })
      .catch(function (e) {
        // Sin red, sin respuesta, tiempo agotado o error del servidor: se guarda en el teléfono y se reenvía después
        var fallo = clasificarFallo(e);
        ultimoFallo = fallo.cod;
        encolar(reg);
        guardarHist(reg, 'pendiente');
        mostrarExito(reg, 'pendiente', fallo);
      })
      .then(function () {
        setCargando(false);
        pintarPendientes();
        pintarHist();
      });
  }

  // "Nuevo registro": conserva país y SVC; limpia el nodo y las sacas para el siguiente nodo
  function aoNovo() {
    el.sacas.value = '';
    nodoActual = null;
    limpiarErrores();
    pintarNodoBtn();
    pintarNodoNota();
    pintarNodoInfo();
    actualizarAvance();
    el.exitoCard.hidden = true;
    el.formCard.hidden = false;
    window.scrollTo(0, 0);
    if (!el.nodoBtn.disabled) el.nodoBtn.focus();
  }

  // ---------------------------------------------------------------------
  // Estado de conexión
  // ---------------------------------------------------------------------
  function pintarConexion() {
    var offline = navigator.onLine === false;
    el.estadoTag.hidden = !(offline || DEMO);
    el.estadoTag.className = 'tag ' + (offline ? 'tag--offline' : 'tag--demo');
    el.estadoTag.textContent = offline ? t('offline') : t('demo');
    el.footer.textContent = DEMO ? t('footer_demo') : t('footer_live');
  }

  // ---------------------------------------------------------------------
  function iniciar() {
    pintarPaises();

    // 1) Idioma provisional según el teléfono, 2) país y su idioma
    aplicarIdioma(idiomaPorDefecto());
    var guardado = leer(LS.pais, '');
    var elegido = paisDeLaUrl();
    if (elegido) {
      seleccionarPais(elegido, true);
    } else if (typeof guardado === 'string' && paisPorCodigo(guardado)) {
      seleccionarPais(guardado, false);
    } else {
      var detectado = paisPorZonaHoraria() || paisPorIdiomaDelTelefono();
      if (detectado) { paisAuto = true; seleccionarPais(detectado, false); }
    }
    var sGuardado = leer(LS.site, null);
    if (sGuardado && sGuardado.pais === paisActual && (sitesPorPais[paisActual] || []).indexOf(sGuardado.site) !== -1) {
      siteActual = sGuardado.site;
    }
    pintarSiteBtn();
    pintarPaisNota();
    pintarNodoBtn();

    el.paises.addEventListener('change', function (ev) {
      if (ev.target && ev.target.name === 'pais') {
        seleccionarPais(ev.target.value, true);
        el.paises.classList.remove('is-invalid');
        limpiarCampo(null, el.paisErr);
      }
    });
    el.siteBtn.addEventListener('click', function () { abrirSheet('site'); });
    el.nodoBtn.addEventListener('click', function () {
      if (nodosEstado === 'error') cargarNodos(siteActual);
      else abrirSheet('nodo');
    });
    el.sheetClose.addEventListener('click', cerrarSheet);
    el.sheetBackdrop.addEventListener('click', cerrarSheet);
    el.sheet.addEventListener('keydown', atraparTab);
    el.siteBuscar.addEventListener('input', pintarLista);
    el.siteLista.addEventListener('click', function (ev) {
      var b = ev.target.closest ? ev.target.closest('.sheet__item') : null;
      if (!b) return;
      var manual = b.getAttribute('data-manual');
      if (manual) { elegirNodo({ id: manual, d1: null, t3: null }); return; }
      var valor = b.getAttribute('data-valor');
      if (sheetModo === 'site') { elegirSite(valor); return; }
      for (var i = 0; i < nodosLista.length; i++) {
        if (nodosLista[i].id === valor) { elegirNodo({ id: valor, d1: Number(nodosLista[i].d1) || 0, t3: Number(nodosLista[i].t3) || 0 }); return; }
      }
    });

    el.sacas.addEventListener('input', function () { limpiarCampo(el.sacas, el.sacasErr); soloDigitos(); });
    el.menos.addEventListener('click', function () { sumarSacas(-1); });
    el.mas.addEventListener('click', function () { sumarSacas(1); });
    el.form.addEventListener('submit', aoSubmit);
    el.novo.addEventListener('click', aoNovo);
    el.reintentar.addEventListener('click', vaciarCola);
    window.addEventListener('online', function () { pintarConexion(); vaciarCola(); if (siteActual && nodosEstado === 'error') cargarNodos(siteActual); });
    window.addEventListener('offline', pintarConexion);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) { depurarHist(); pintarHist(); } });

    pintarConexion();
    pintarPendientes();
    depurarHist();
    pintarHist();
    actualizarAvance();

    if (DEMO && !DEMO_PERMITIDO) {
      el.erroCatalogo.textContent = 'Falta configurar la app: pega la URL del servidor (Apps Script) en API_URL, dentro de app.js.';
      el.erroCatalogo.hidden = false;
      el.enviar.disabled = true;
      return;
    }

    if (siteActual) cargarNodos(siteActual);

    cargarCatalogo().then(function (cat) {
      pintarCatalogo(cat);
      el.enviar.disabled = false;
      actualizarAvance();
      vaciarCola();
    });
  }

  iniciar();
})();
