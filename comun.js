/* =====================================================================
   comun.js · Unidad de Estadística · Hospital San Luis de Buin-Paine
   Código compartido por TODAS las páginas:
     · encabezado con menú y pie de página (se escriben aquí una sola vez)
     · búsqueda y lectura de los REM (SA y SBS) y de las bases Excel
     · formatos de números, animaciones, mini-gráficos y variaciones
     · botones "Descargar Excel" e "Imprimir"
   Para agregar una página al menú: edita la lista MENU_PRODUCCION.
   ===================================================================== */
(function () {
  "use strict";

  /* ---------------- CONFIGURACIÓN GENERAL ---------------- */
  const SITIO = {
    anioDesde: 2024,      // primer año en que se buscan REM
    /* dónde puede estar la carpeta REM respecto de la página */
    prefijos: ["", "../", "PAGINA ESTADÍSTICA/", "PAGINA ESTADISTICA/"],
    rem: {
      /* REM serie A: REM/<año>/REM_SA/SA_01.xlsm … */
      SA: {
        carpetas: ["REM_SA", "REM_A", "REM SA", "REM A"],
        nombres: [mm => `SA_${mm}`],
        extensiones: [".xlsm", ".xlsx", ".xls"]
      },
      /* REM serie BS: REM/<año>/REM_BS/SBS_01_26.xlsm … */
      SBS: {
        carpetas: ["REM_BS", "REM BS", "REM_SBS", "REM_B", "SBS", "REM SBS", "REM B", "REM_SA", "REM_A"],
        nombres: [(mm, aa) => `SBS_${mm}_${aa}`, mm => `SBS_${mm}`],
        extensiones: [".xlsm", ".xlsx"]
      }
    },
    contacto: { correo: "Mauricio.laram@redsalud.gob.cl", anexo: "264881", telefono: "2-2 5767881", tel: "+56225767881" }
  };

  /* Páginas del menú "Producción" (clave = data-pagina del <body>) */
  const MENU_PRODUCCION = [
    { clave: "urgencia",   texto: "Urgencia",              href: "urgencia.html",            icono: "urgencia" },
    { clave: "pabellon",   texto: "Pabellón",              href: "pabellon.html",            icono: "pabellon" },
    { clave: "cerrada",    texto: "Atención Cerrada",      href: "atencion-cerrada.html",    icono: "cerrada" },
    { clave: "medica",     texto: "Ambulatorio Médica",    href: "produccion-medica.html",   icono: "medica" },
    { clave: "nomedica",   texto: "Ambulatorio no Médica", href: "produccion-no-medica.html", icono: "nomedica", pronto: true },
    { clave: "dental",     texto: "Odontología",           href: "produccion-dental.html",   icono: "dental" },
    { clave: "laboratorio", texto: "Laboratorio",          href: "laboratorio.html",         icono: "laboratorio" },
    { clave: "imagenologia", texto: "Imagenología",        href: "imagenologia.html",        icono: "imagenologia" }
  ];

  /* Íconos (trazos simples, heredan el color del texto) */
  const ICONOS = {
    reloj: '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
    urgencia: '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>',
    pabellon: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.1 15.9M14.5 14.5 20 20M8.1 8.1 12 12"/>',
    cerrada: '<path d="M2 4v16M2 8h18a2 2 0 0 1 2 2v10M2 17h20M6 8v9"/>',
    medica: '<path d="M4.8 2.3A.3.3 0 1 0 5 2H4a2 2 0 0 0-2 2v5a6 6 0 0 0 12 0V4a2 2 0 0 0-2-2h-1a.2.2 0 1 0 .3.3"/><path d="M8 15v1a6 6 0 0 0 12 0v-4"/><circle cx="20" cy="10" r="2"/>',
    nomedica: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    dental: '<path d="M12 5.5C10.5 4 8 3 6 4S3 8 4 11s1.5 9 3.5 9 2-5 4.5-5 2.5 5 4.5 5S19 14 20 11s0-6-2-7-4.5 0-6 1.5z"/>',
    laboratorio: '<path d="M9 2h6M10 2v7L4.5 19a2 2 0 0 0 1.8 3h11.4a2 2 0 0 0 1.8-3L14 9V2M7 15h10"/>',
    imagenologia: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="12" cy="11" r="4"/><path d="M12 7v8M8 11h8M7 19h10"/>',
    documento: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8"/>',
    descargar: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>',
    imprimir: '<path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v8H6z"/>',
    indicador: '<path d="M3 3v18h18M7 16v-5M12 16V8M17 16V7"/>',
    flecha: '<path d="M5 12h14M13 6l6 6-6 6"/>',
    abajo: '<path d="m6 9 6 6 6-6"/>',
    menu: '<path d="M3 6h18M3 12h18M3 18h18"/>',
    cerrar: '<path d="M18 6 6 18M6 6l12 12"/>',
    buscar: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    ver: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
    correo: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 6-10 7L2 6"/>',
    telefono: '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/>'
  };
  const icono = (nombre, tam = 22) =>
    `<svg width="${tam}" height="${tam}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ` +
    `stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONOS[nombre] || ""}</svg>`;

  /* ---------------- FORMATOS Y UTILIDADES ---------------- */
  const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const MESES_CORTOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const COLOR = { marino: "#0a132d", azul: "#0f69b4", oscuro: "#0a4f8a", rojo: "#eb3c46", gris: "#5f6b7a", grisClaro: "#8a9097", claro: "#e8f1fa", linea: "#dde3ea" };

  const $ = id => document.getElementById(id);
  /* mayúsculas, sin tildes y sin espacios de más: para comparar textos del REM */
  const norm = v => String(v ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
  const n = v => Number(v) || 0;
  const num = v => Number(v || 0).toLocaleString("es-CL");
  const dec = (v, d = 1) => v == null || !isFinite(v) ? "–" : Number(v).toLocaleString("es-CL", { minimumFractionDigits: d, maximumFractionDigits: d });
  const pct = (a, b) => b ? dec(a / b * 100, 1) + "%" : "–";
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const el = x => typeof x === "string" ? $(x) : x;
  const reducirMovimiento = () => window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* Cuenta hacia arriba hasta el valor (animación de los KPI) */
  function animar(destino, valor, formato = num) {
    const e = el(destino);
    if (!e) return;
    if (valor == null || !isFinite(valor)) { e.textContent = "–"; return; }
    const inicio = Number(e.dataset.valor || 0);
    e.dataset.valor = valor;
    if (reducirMovimiento() || document.hidden) { e.textContent = formato(valor); return; }
    const t0 = performance.now(), dur = 700;
    cancelAnimationFrame(e._raf);
    const paso = t => {
      const p = Math.min(1, (t - t0) / dur), suave = 1 - Math.pow(1 - p, 3);
      const actual = inicio + (valor - inicio) * suave;
      e.textContent = formato(p < 1 && Number.isInteger(valor) ? Math.round(actual) : (p < 1 ? actual : valor));
      if (p < 1) e._raf = requestAnimationFrame(paso);
    };
    e._raf = requestAnimationFrame(paso);
  }

  /* Variación respecto del período anterior.
     sentido: 1 = subir es bueno (verde), -1 = bajar es bueno, 0 = sin color.
     tipo: "%" (variación porcentual) o "pt" (diferencia en puntos, para indicadores en %) */
  function variacion(destino, actual, anterior, { sentido = 0, tipo = "%", texto = "vs mes anterior", extra = "" } = {}) {
    const e = el(destino);
    if (!e) return;
    if (actual == null || anterior == null || !isFinite(actual) || !isFinite(anterior) || (tipo === "%" && !anterior)) {
      e.innerHTML = extra; return;
    }
    const d = tipo === "%" ? (actual - anterior) / anterior * 100 : actual - anterior;
    const flecha = Math.abs(d) < 0.05 ? "＝" : d > 0 ? "▲" : "▼";
    const clase = !sentido || Math.abs(d) < 0.05 ? "" : (d > 0) === (sentido > 0) ? "bien" : "mal";
    e.innerHTML = `<span class="delta ${clase}">${flecha} ${dec(Math.abs(d), 1)}${tipo === "%" ? "%" : " pt"}</span>${esc(texto)}${extra}`;
  }

  /* Mini-gráfico de tendencia dentro de un <svg class="spark"> */
  function sparkline(destino, valores, color = COLOR.azul) {
    const e = el(destino);
    if (!e) return;
    const v = (valores || []).map(x => x == null ? null : Number(x));
    const ok = v.filter(x => x != null && isFinite(x));
    if (ok.length < 2) { e.innerHTML = ""; return; }
    const W = 96, H = 28, lo = Math.min(...ok), hi = Math.max(...ok);
    e.setAttribute("viewBox", `0 0 ${W} ${H}`);
    e.setAttribute("aria-hidden", "true");
    const pts = v.map((x, i) => x == null ? null : [i * W / (v.length - 1), H - 3 - (x - lo) / ((hi - lo) || 1) * (H - 6)]).filter(Boolean);
    const ult = pts[pts.length - 1];
    e.innerHTML = `<polyline points="${pts.map(p => p.map(c => c.toFixed(1)).join(",")).join(" ")}" fill="none" stroke="${color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>` +
                  `<circle cx="${ult[0].toFixed(1)}" cy="${ult[1].toFixed(1)}" r="3" fill="${color}"/>`;
  }

  /* ---------------- AVISOS Y ESTADO DE CARGA ---------------- */
  function aviso(html, { error = false, progreso = false } = {}) {
    const e = $("aviso");
    if (!e) return;
    e.hidden = false;
    e.classList.toggle("error", error);
    e.innerHTML = html + (progreso ? '<div class="progreso"><div></div></div>' : "");
  }
  function ocultarAviso() { const e = $("aviso"); if (e) e.hidden = true; }
  function listo(texto) {
    document.body.classList.remove("cargando");
    if (texto && $("actualizado")) $("actualizado").textContent = texto;
  }
  /* Texto del chip de la cabecera: "Datos a octubre 2026" */
  const textoActualizado = (anio, mes) => `Datos a ${MESES[mes - 1].toLowerCase()} ${anio}`;
  /* Último año y mes presentes en una lista [{anio, mes}] */
  const ultimoPeriodo = lista => lista.reduce((u, d) => !u || d.anio * 100 + d.mes > u.anio * 100 + u.mes ? { anio: d.anio, mes: d.mes } : u, null);

  function avisoArchivoLocal() {
    if (location.protocol !== "file:") return false;
    aviso("La página está abierta con doble clic (la dirección empieza con <b>file://</b>). " +
      "Así el navegador bloquea la lectura de los archivos Excel.<br><small>Ábrela desde el servidor de la página " +
      "(la dirección debe empezar con http:// o https://).</small>", { error: true });
    document.body.classList.remove("cargando");
    if ($("actualizado")) $("actualizado").textContent = "Sin datos";
    return true;
  }

  /* ---------------- LECTURA DE ARCHIVOS ---------------- */
  async function existe(url) {
    try {
      const r = await fetch(encodeURI(url), { method: "HEAD", cache: "no-cache" });
      if (r.ok) return true;
      if (r.status === 405 || r.status === 501) return (await fetch(encodeURI(url), { cache: "no-cache" })).ok;
    } catch (e) { /* sin conexión o ruta inválida */ }
    return false;
  }

  /* Descarga un Excel. cache "no-cache": el navegador pregunta al servidor si cambió
     y, si no cambió, usa la copia guardada (mucho más rápido al volver a entrar). */
  async function descargarLibro(url, opciones = {}) {
    const r = await fetch(encodeURI(url), { cache: "no-cache" });
    if (!r.ok) return null;
    return XLSX.read(await r.arrayBuffer(), { type: "array", ...opciones });
  }

  /* Busca un Excel probando rutas y prefijos; devuelve { wb, ruta } o null */
  async function abrirExcel(rutas, opciones = {}) {
    for (const pre of SITIO.prefijos) for (const ruta of rutas) {
      try {
        const wb = await descargarLibro(pre + ruta, opciones);
        if (wb) return { wb, ruta: pre + ruta };
      } catch (e) { console.warn("No se pudo leer", pre + ruta, e); }
    }
    return null;
  }

  const aniosBuscados = () => { const a = [], hoy = new Date().getFullYear(); for (let x = hoy; x >= SITIO.anioDesde; x--) a.push(x); return a; };

  /* Encuentra la carpeta y el nombre de los REM (SA o SBS).
     Recuerda la ruta encontrada en el navegador para no volver a buscarla. */
  async function buscarRutaREM(tipo) {
    const cfg = SITIO.rem[tipo], anios = aniosBuscados();
    const candidatos = [];
    for (const pre of SITIO.prefijos) for (const car of cfg.carpetas) for (let k = 0; k < cfg.nombres.length; k++) for (const ext of cfg.extensiones)
      candidatos.push({ pre, car, k, ext });
    const armar = c => (anio, mes) => `${c.pre}REM/${anio}/${c.car}/${cfg.nombres[c.k](String(mes).padStart(2, "0"), String(anio).slice(-2))}${c.ext}`;
    const clave = "estadistica.rutaREM." + tipo;
    const tiene = async (c, listaAnios) => {
      const r = armar(c);
      const pruebas = [];
      for (const a of listaAnios) for (let m = 1; m <= 12; m++) pruebas.push(existe(r(a, m)));
      return (await Promise.all(pruebas)).some(Boolean);
    };
    /* 1) la ruta que funcionó la última vez */
    try {
      const g = JSON.parse(localStorage.getItem(clave) || "null");
      if (g && candidatos.some(c => JSON.stringify(c) === JSON.stringify(g)) && await tiene(g, anios)) return armar(g);
    } catch (e) { /* almacenamiento no disponible */ }
    /* 2) búsqueda: primero el año actual y el anterior, luego el resto; de a 6 rutas en paralelo */
    for (const listaAnios of [anios.slice(0, 2), anios.slice(2)]) {
      if (!listaAnios.length) continue;
      for (let i = 0; i < candidatos.length; i += 6) {
        const grupo = candidatos.slice(i, i + 6);
        const res = await Promise.all(grupo.map(c => tiene(c, listaAnios)));
        const k = res.indexOf(true);
        if (k >= 0) {
          try { localStorage.setItem(clave, JSON.stringify(grupo[k])); } catch (e) { /* nada */ }
          return armar(grupo[k]);
        }
      }
    }
    return null;
  }

  /* Lee una hoja de todos los REM disponibles.
     leer(libro, {anio, mes}) devuelve un objeto con los datos del mes (o null).
     Resultado: { datos: [{anio, mes, ...}], ruta } */
  async function cargarREM(tipo, hoja, leer) {
    const ruta = await buscarRutaREM(tipo);
    if (!ruta) return { datos: [], ruta: null };
    const tareas = [];
    let leidos = 0;
    for (const anio of aniosBuscados()) for (let mes = 1; mes <= 12; mes++) {
      tareas.push((async () => {
        try {
          const wb = await descargarLibro(ruta(anio, mes), { sheets: Array.isArray(hoja) ? hoja : [hoja] });
          if (!wb) return null;
          leidos++;
          aviso(`Leyendo REM ${tipo}… ${leidos} ${leidos === 1 ? "archivo" : "archivos"}`, { progreso: true });
          const d = leer(wb, { anio, mes });
          return d ? { anio, mes, ...d } : null;
        } catch (e) { console.warn("REM", tipo, anio, mes, e); return null; }
      })());
    }
    const datos = (await Promise.all(tareas)).filter(Boolean).sort((a, b) => a.anio - b.anio || a.mes - b.mes);
    return { datos, ruta };
  }

  /* Año y mes de una fecha de Excel (número de serie o texto) */
  function anioMes(v) {
    if (v instanceof Date) return { anio: v.getFullYear(), mes: v.getMonth() + 1 };
    if (typeof v === "number") { const d = XLSX.SSF.parse_date_code(v); return d ? { anio: d.y, mes: d.m } : null; }
    const t = String(v ?? "").trim();
    let m = t.match(/^(\d{4})[-\/](\d{1,2})/);
    if (m) return { anio: +m[1], mes: +m[2] };
    m = t.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (m) return { anio: +m[3], mes: +m[2] };
    return null;
  }

  /* ---------------- FILTROS AÑO / MES (comunes a las páginas de producción) ---------------- */
  /* periodos: lista [{anio, mes}] con datos. alCambiar: función que redibuja la página.
     Por defecto queda elegido el ÚLTIMO MES con datos (así se ve la variación vs el mes anterior). */
  function filtrosPeriodo(periodos, alCambiar, { mesPorDefecto = "ultimo" } = {}) {
    const fA = $("fAnio"), fM = $("fMes");
    const anios = [...new Set(periodos.map(d => d.anio))].sort((a, b) => b - a);
    fA.innerHTML = anios.map(a => `<option value="${a}">${a}</option>`).join("");
    fA.value = anios[0];
    const llenarMeses = (elegir) => {
      const previo = fM.value;
      const meses = [...new Set(periodos.filter(d => d.anio == fA.value).map(d => d.mes))].sort((a, b) => a - b);
      fM.innerHTML = '<option value="">Todos (acumulado)</option>' + meses.map(m => `<option value="${m}">${MESES[m - 1]}</option>`).join("");
      if (elegir === "ultimo" && meses.length) fM.value = meses[meses.length - 1];
      else if (meses.includes(Number(previo))) fM.value = previo;
    };
    llenarMeses(mesPorDefecto);
    fA.onchange = () => { llenarMeses(); alCambiar(); };
    fM.onchange = alCambiar;
  }
  /* Etiqueta del período elegido: "Octubre 2026" o "Acumulado 2026" */
  const etiquetaPeriodo = () => { const a = $("fAnio")?.value, m = $("fMes")?.value; return m ? `${MESES[m - 1]} ${a}` : `Acumulado ${a}`; };
  /* Mes anterior al elegido (para la variación) */
  const mesAnterior = (anio, mes) => mes > 1 ? { anio: Number(anio), mes: mes - 1 } : { anio: Number(anio) - 1, mes: 12 };

  /* ---------------- GRÁFICOS (Chart.js) ---------------- */
  const graf = {};
  if (window.Chart) {
    if (window.ChartDataLabels) Chart.register(ChartDataLabels);
    Chart.defaults.font.family = '"Roboto", Arial, sans-serif';
    Chart.defaults.color = COLOR.gris;
    Chart.defaults.maintainAspectRatio = false;
    Chart.defaults.resizeDelay = 100;
    Chart.defaults.animation.duration = reducirMovimiento() ? 0 : 700;
    const tam = (ctx, chico, grande) => { const w = ctx.chart.width || 800; return w < 420 ? chico : w < 700 ? Math.round((chico + grande) / 2) : grande; };
    graf.tam = tam;
    graf.fuenteEje = ctx => ({ size: tam(ctx, 10, 13), weight: "500" });
    graf.fuenteValor = ctx => ({ size: tam(ctx, 9, 12), weight: "600" });
    graf.tooltip = {
      backgroundColor: COLOR.marino, titleColor: "#fff", bodyColor: "#e8f1fa", padding: 10, cornerRadius: 8,
      callbacks: {
        title: items => String(items[0].label).replace(/,/g, " "),
        label: ctx => ` ${ctx.dataset.label ? ctx.dataset.label + ": " : ""}${num(ctx.raw)}`
      }
    };
    graf.leyenda = { position: "top", align: "center", labels: { usePointStyle: true, pointStyle: "rectRounded", color: COLOR.marino, font: graf.fuenteEje } };
    graf.ejeValores = { beginAtZero: true, grace: "14%", ticks: { callback: v => num(v), font: graf.fuenteEje, precision: 0 },
                        grid: { color: COLOR.linea }, border: { display: false } };
    graf.etiquetas = { anchor: "end", align: "end", offset: 2, clamp: true, color: COLOR.marino, font: graf.fuenteValor,
                       formatter: v => num(v), display: ctx => Number(ctx.dataset.data[ctx.dataIndex]) > 0 };
  }
  /* Vuelve a medir los gráficos cuando cambia el tamaño de la página */
  function vigilarGraficos(graficos) {
    const ajustar = () => requestAnimationFrame(() => graficos.forEach(g => g && g.resize()));
    document.querySelectorAll(".chart-box").forEach(c => new ResizeObserver(ajustar).observe(c));
    window.addEventListener("load", ajustar);
    if (document.fonts) document.fonts.ready.then(ajustar);
    document.querySelectorAll("details").forEach(d => d.addEventListener("toggle", ajustar));
    return ajustar;
  }
  /* Etiquetas largas en varias líneas */
  function partir(txt, max) {
    txt = String(txt);
    if (txt.length <= max) return txt;
    const lineas = [""];
    for (const p of txt.split(" ")) {
      const i = lineas.length - 1;
      if ((lineas[i] + " " + p).trim().length > max && lineas[i]) lineas.push(p); else lineas[i] = (lineas[i] + " " + p).trim();
    }
    return lineas;
  }

  /* ---------------- DESCARGAR EXCEL E IMPRIMIR ---------------- */
  /* Convierte el texto de una celda ("1.234", "45,6%") en número para Excel */
  function celdaExcel(texto) {
    const t = texto.replace(/\s+/g, " ").trim();
    if (/^-?\d{1,3}(\.\d{3})*(,\d+)?%$/.test(t) || /^-?\d+(,\d+)?%$/.test(t)) return { v: Number(t.slice(0, -1).replace(/\./g, "").replace(",", ".")) / 100, t: "n", z: "0.0%" };
    if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(t) || /^-?\d+(,\d+)?$/.test(t)) return { v: Number(t.replace(/\./g, "").replace(",", ".")), t: "n" };
    return { v: t, t: "s" };
  }
  function exportarExcel() {
    if (!window.XLSX) return;
    const tablas = [...document.querySelectorAll("main table")].filter(t => t.rows.length);
    if (!tablas.length) { alert("Esta página no tiene tablas para descargar."); return; }
    const libro = XLSX.utils.book_new(), usados = new Set();
    for (const tabla of tablas) {
      const filas = [];
      for (const tr of tabla.rows) {
        const fila = [];
        for (const c of tr.cells) {
          fila.push(celdaExcel(c.innerText || c.textContent || ""));
          for (let k = 1; k < (c.colSpan || 1); k++) fila.push({ v: "", t: "s" });
        }
        filas.push(fila);
      }
      const hoja = XLSX.utils.aoa_to_sheet(filas.map(f => f.map(c => c.v)));
      filas.forEach((f, r) => f.forEach((c, k) => { if (c.z) { const ref = XLSX.utils.encode_cell({ r, c: k }); if (hoja[ref]) hoja[ref].z = c.z; } }));
      hoja["!cols"] = filas[0].map((_, k) => ({ wch: k ? 14 : 34 }));
      let nombre = (tabla.closest(".tarjeta")?.querySelector("h2")?.textContent || "Tabla").replace(/[\\\/\?\*\[\]:]/g, " ").trim().slice(0, 28) || "Tabla";
      let final = nombre, i = 2;
      while (usados.has(final)) final = `${nombre.slice(0, 26)} ${i++}`;
      usados.add(final);
      XLSX.utils.book_append_sheet(libro, hoja, final);
    }
    const pagina = document.body.dataset.pagina || "estadistica";
    const periodo = etiquetaPeriodo ? (($("fAnio") && $("fAnio").value) ? "_" + etiquetaPeriodo().replace(/\s+/g, "_") : "") : "";
    XLSX.writeFile(libro, `${pagina}${periodo}.xlsx`);
  }
  function imprimir() {
    const cerrados = [...document.querySelectorAll("details:not([open])")];
    cerrados.forEach(d => d.open = true);
    setTimeout(() => {
      window.print();
      cerrados.forEach(d => d.open = false);
    }, 300);
  }

  /* ---------------- ENCABEZADO, MENÚ Y PIE ---------------- */
  /* ---------- Fecha y hora actual (se actualiza sola en todas las páginas) ---------- */
  function mayuscula(t) { return t.charAt(0).toUpperCase() + t.slice(1); }
  function pintarFechaHora() {
    const caja = $("fecha-hora");
    if (!caja) return;
    const ahora = new Date();
    const larga = mayuscula(ahora.toLocaleDateString("es-CL", { weekday: "long", day: "numeric", month: "long", year: "numeric" }));
    const corta = ahora.toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric" });
    const hora = ahora.toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit", hour12: false });
    caja.dateTime = ahora.toISOString();
    caja.innerHTML = `<span class="fecha-larga">${larga}</span><span class="fecha-corta">${corta}</span> · <span class="hora">${hora} hrs</span>`;
  }
  function iniciarReloj() {
    pintarFechaHora();
    /* se sincroniza con el cambio de minuto y luego se repite cada minuto */
    setTimeout(() => { pintarFechaHora(); setInterval(pintarFechaHora, 60000); }, 60000 - Date.now() % 60000 + 50);
  }

  function construirEncabezado() {
    const pagina = document.body.dataset.pagina || "";
    const enProduccion = MENU_PRODUCCION.some(p => p.clave === pagina);
    const sub = MENU_PRODUCCION.map(p =>
      `<a href="${p.href}" class="${p.clave === pagina ? "activo" : ""}"${p.clave === pagina ? ' aria-current="page"' : ""}>${icono(p.icono, 18)}${p.texto}${p.pronto ? "<small>Pronto</small>" : ""}</a>`).join("");
    const html = `
      <div class="barra-gobierno"><span class="azul"></span><span class="rojo"></span></div>
      <div class="barra-fecha"><div class="contenedor">${icono("reloj", 15)}<time id="fecha-hora"></time></div></div>
      <header class="encabezado">
        <div class="contenedor encabezado-fila">
          <a href="index.html" class="marca">
            <div class="marca-texto">
              <span class="marca-red">Servicio de Salud Metropolitano Sur</span>
              <strong>Hospital San Luis de Buin-Paine</strong>
              <span class="marca-unidad">Unidad de Estadística</span>
            </div>
          </a>
          <button class="menu-boton" type="button" aria-expanded="false" aria-controls="menu-principal" aria-label="Abrir menú">${icono("menu", 22)}</button>
          <nav class="menu" id="menu-principal" aria-label="Menú principal">
            <a href="index.html" class="${pagina === "inicio" ? "activo" : ""}">Inicio</a>
            <div class="menu-grupo">
              <button class="menu-desplegar ${enProduccion ? "activo" : ""}" type="button" aria-expanded="false">Producción ${icono("abajo", 16)}</button>
              <div class="submenu">${sub}</div>
            </div>
            <a href="manuales.html" class="${pagina === "manuales" ? "activo" : ""}">Manuales</a>
            <a href="index.html#contacto">Contacto</a>
          </nav>
        </div>
      </header>`;
    document.body.insertAdjacentHTML("afterbegin", html);

    /* menú desplegable y menú de celular */
    const grupo = document.querySelector(".menu-grupo"), boton = grupo.querySelector(".menu-desplegar");
    const nav = $("menu-principal"), hamburguesa = document.querySelector(".menu-boton");
    const abrirGrupo = si => { grupo.classList.toggle("abierto", si); boton.setAttribute("aria-expanded", si); };
    boton.addEventListener("click", e => { e.stopPropagation(); abrirGrupo(!grupo.classList.contains("abierto")); });
    if (matchMedia("(hover: hover) and (min-width: 901px)").matches) {
      grupo.addEventListener("mouseenter", () => abrirGrupo(true));
      grupo.addEventListener("mouseleave", () => abrirGrupo(false));
    }
    hamburguesa.addEventListener("click", e => {
      e.stopPropagation();
      const si = !nav.classList.contains("abierto");
      nav.classList.toggle("abierto", si);
      hamburguesa.setAttribute("aria-expanded", si);
      hamburguesa.innerHTML = icono(si ? "cerrar" : "menu", 22);
      if (si && enProduccion) abrirGrupo(true);
    });
    document.addEventListener("click", e => {
      if (!grupo.contains(e.target)) abrirGrupo(false);
      if (!nav.contains(e.target) && !hamburguesa.contains(e.target)) { nav.classList.remove("abierto"); hamburguesa.setAttribute("aria-expanded", false); hamburguesa.innerHTML = icono("menu", 22); }
    });
    document.addEventListener("keydown", e => { if (e.key === "Escape") { abrirGrupo(false); nav.classList.remove("abierto"); } });
  }

  function construirPie() {
    const c = SITIO.contacto;
    document.body.insertAdjacentHTML("beforeend", `
      <footer class="pie">
        <div class="contenedor pie-columnas">
          <div><strong>Hospital San Luis de Buin-Paine</strong>Unidad de Estadística<br>Servicio de Salud Metropolitano Sur</div>
          <div><strong>Contacto</strong><a href="mailto:${c.correo}">${c.correo}</a><br>Anexo ${c.anexo} · <a href="tel:${c.tel}">${c.telefono}</a></div>
          <div><strong>Fuentes</strong>REM Serie A y BS del Ministerio de Salud y bases locales de urgencia, pabellón y atención cerrada.</div>
        </div>
        <div class="contenedor pie-nota">Datos agregados. No contiene información de pacientes. · © ${new Date().getFullYear()} Unidad de Estadística</div>
      </footer>`);
  }

  /* Íconos escritos como <span data-icono="urgencia"></span> en el HTML */
  function pintarIconos() {
    document.querySelectorAll("[data-icono]").forEach(e => { e.innerHTML = icono(e.dataset.icono, Number(e.dataset.tam) || 26); });
  }

  /* Las tarjetas aparecen suavemente al bajar por la página */
  function revelarAlBajar() {
    if (reducirMovimiento() || !("IntersectionObserver" in window)) return;
    const io = new IntersectionObserver(entradas => entradas.forEach(en => {
      if (en.isIntersecting) {
        const e = en.target;
        e.classList.add("visible");
        io.unobserve(e);
        /* al terminar, se quitan las clases para que el efecto al pasar el mouse siga funcionando */
        setTimeout(() => e.classList.remove("revelar", "visible"), 600);
      }
    }), { rootMargin: "0px 0px -40px 0px" });
    document.querySelectorAll("main .tarjeta, main .modulo").forEach(e => {
      const r = e.getBoundingClientRect();
      if (r.top < innerHeight) return;            // lo que ya se ve al cargar no se anima
      e.classList.add("revelar");
      io.observe(e);
    });
  }

  function iniciar() {
    construirEncabezado();
    iniciarReloj();
    construirPie();
    pintarIconos();
    document.querySelectorAll("[data-exportar]").forEach(b => b.addEventListener("click", exportarExcel));
    document.querySelectorAll("[data-imprimir]").forEach(b => b.addEventListener("click", imprimir));
    window.addEventListener("load", revelarAlBajar);
  }

  /* ---------------- LO QUE QUEDA DISPONIBLE PARA LAS PÁGINAS ---------------- */
  window.SITIO = SITIO;
  window.U = {
    MESES, MESES_CORTOS, COLOR, ICONOS, icono,
    $, norm, n, num, dec, pct, esc,
    animar, variacion, sparkline, partir,
    aviso, ocultarAviso, listo, textoActualizado, ultimoPeriodo, avisoArchivoLocal,
    existe, descargarLibro, abrirExcel, buscarRutaREM, cargarREM, anioMes,
    filtrosPeriodo, etiquetaPeriodo, mesAnterior,
    graf, vigilarGraficos
  };

  iniciar();
})();
