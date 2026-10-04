'use strict';
/* ControlPlus CCO — interfaz. Se comunica con Kotlin/Python mediante window.Android. */

const VERSION = '1.0.0';
const $ = (s, r = document) => r.querySelector(s);
function h(tag, props = {}, ...hijos) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (k === 'class') e.className = v;
    else if (k === 'html') e.innerHTML = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  hijos.flat().forEach(c => { if (c != null && c !== false) e.append(c.nodeType ? c : document.createTextNode(c)); });
  return e;
}
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const pad = n => String(n).padStart(2, '0');
const horaAhora = () => { const d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); };
const fechaAhora = () => { const d = new Date(); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${horaAhora()}`; };
const kb = b => b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
function toast(t) { const e = $('#toast'); e.textContent = t; e.classList.add('ver'); clearTimeout(toast._t); toast._t = setTimeout(() => e.classList.remove('ver'), 2200); }

/* ── Puente nativo (con vista previa solo para navegador de escritorio) ─────── */
const PREVIEW = typeof window.Android === 'undefined';
if (PREVIEW) {
  // SOLO vista previa del diseño fuera de la APK. En la APK real se usa el puente Kotlin→Python.
  window.Android = {
    elegirArchivos: (id, multi) => setTimeout(() => window.__nativo(id, JSON.stringify({ ok: true, archivos: [{ nombre: 'ejemplo_vista_previa.xls', ruta: '/preview/ejemplo.xls', bytes: 94059 }] })), 350),
    ejecutar: (id, her) => setTimeout(() => window.__nativo(id, JSON.stringify({ ok: true, logs: ['(vista previa)'], mensajes: [
      { titulo: 'Ejemplo 1', texto: 'Buenas Tardes\n\n* Eje *Blv* Sub Eje *Acevedo*\n* Unidad *0645* - Vin *0684*\n* CTS: *Johan S.*\n\n_Incidencia:_ Exceso de velocidad\n_Incumplimiento:_ *93 km/h* en *autopista*' }] })), 1800),
    unidadConocida: (id, u) => setTimeout(() => window.__nativo(id, JSON.stringify({ alias: u.toUpperCase(), info: ['Blv', 'Acevedo', '0684'] })), 200),
    copiar: t => navigator.clipboard && navigator.clipboard.writeText(t), compartirTexto: t => alert('Compartir:\n' + t),
    compartirArchivo: r => alert('Compartir archivo ' + r), guardarArchivo: r => JSON.stringify({ ok: true, nombre: r.split('/').pop(), carpeta: 'Descargas/ControlPlus' }),
    aviso: t => toast(t), salir: () => { }
  };
  setTimeout(() => window.__nativo(0, JSON.stringify({ version: VERSION, compat: { tkinter: 'vista previa', calamine: 'vista previa', lxml: 'vista previa' } })), 1200);
}
let _id = 1; const _pend = {};
window.__nativo = (id, json) => {
  let d; try { d = JSON.parse(json); } catch (e) { d = { ok: false, error: 'Respuesta inválida del motor.' }; }
  if (id === 0) { motor.estado = d.error ? 'error' : 'listo'; motor.info = d; pintarMotor(); return; }
  const r = _pend[id]; if (r) { delete _pend[id]; r(d); }
};
const llamar = (fn, ...a) => new Promise(res => { const id = _id++; _pend[id] = res; window.Android[fn](id, ...a); });
const motor = { estado: 'espera', info: null };
function pintarMotor() {
  const c = $('#chip-motor'); if (!c) return;
  const t = { espera: 'Preparando motor…', listo: 'Motor listo', error: 'Motor con error (ver ⓘ)' }[motor.estado];
  c.querySelector('.punto').className = 'punto ' + (motor.estado === 'listo' ? '' : motor.estado);
  c.querySelector('span').textContent = t;
}

/* ── Herramientas (flujo equivalente al del bot de Telegram) ───────────────── */
const XLS = ['xls', 'xlsx'];
const T = [
  { id: 'excesos', ic: '🚨', nombre: 'Excesos de Velocidad', corto: 'PDF de alarmas → un mensaje por unidad',
    necesitas: ['PDF «Notificaciones de alarmas» del GTRMax', 'La hora actual (define el saludo)', 'Opcional: Excel de Disponibilidad para completar el CTS'],
    obtienes: 'Un mensaje listo para WhatsApp por cada unidad con exceso, con su mayor velocidad detectada.',
    pasos: [
      { clave: 'pdf', tipo: 'archivo', ext: ['pdf'], titulo: 'PDF de alarmas', ayuda: 'En el GTRMax abre «Notificaciones de alarmas» y expórtalo a PDF. Solo se leen las filas de «Agente de Velocidad».' },
      { clave: 'hora', tipo: 'hora', titulo: '¿Qué hora es ahora?', ayuda: 'Se usa para el saludo: Buenos Días (4:00–11:59), Buenas Tardes (12:00–18:59) o Buenas Noches (19:00–3:59). Acepta 21:15 o 08:02 pm.' },
      { clave: 'excel', tipo: 'archivo', ext: XLS, opcional: true, titulo: 'Excel de Disponibilidad (opcional)', ayuda: 'Si lo adjuntas, el campo CTS se completa con el nombre real. Si no, aparecerá «S.I.».' }] },
  { id: 'sin_info', ic: '📋', nombre: 'Sin Información (Top 5)', corto: 'Las 5 unidades con más km por sub-eje',
    necesitas: ['Reporte diario del GTR (.xls o .xlsx)'],
    obtienes: 'Un mensaje por Eje con las unidades (y su VIN) que más recorrieron y no reportan información.',
    pasos: [{ clave: 'archivo', tipo: 'archivo', ext: XLS, titulo: 'Reporte diario GTR', ayuda: 'Es el «Reporte Operación Flota / Reporte Diario» que descargas del GTR.' }] },
  { id: 'ultimas', ic: '🕒', nombre: 'Últimas UT en Movimiento', corto: 'Unidades con menos de 3 min sin reportar',
    necesitas: ['«Reporte de Estatus Actual» (.xls)', 'La hora para el encabezado', 'Opcional: unidades a obviar'],
    obtienes: 'Un mensaje por grupo (Am, Blv, Met, Ocm, Pz, Serv Cont, Serv Esp) con las unidades en movimiento ahora mismo.',
    pasos: [
      { clave: 'archivo', tipo: 'archivo', ext: XLS, titulo: 'Reporte de Estatus Actual', ayuda: 'Descárgalo del GTR (Estatus Actual). Aunque termine en .xls, es normal: la app sabe leerlo.' },
      { clave: 'hora', tipo: 'hora', titulo: 'Hora del encabezado', ayuda: 'Puedes escribirla en 12 h o 24 h; el mensaje siempre sale en 12 h (ej. 08:00 Pm).' },
      { clave: 'excluir', tipo: 'texto', opcional: true, titulo: '¿Obviar alguna UT? (opcional)', ph: 'Ej: R-066, 0343', ayuda: 'Escribe las unidades separadas por coma. Acepta «r66», «R-66» o «343». Las que no reconozca se ignoran.' }] },
  { id: 'subeje', ic: '🧩', nombre: 'UT en Movimiento + Sub-eje', corto: 'Reporte por Eje y sin información por Sub-eje',
    necesitas: ['Reporte diario del GTR (.xls o .xlsx)', 'La hora del reporte'],
    obtienes: 'El reporte por Eje y un mensaje de «Sin información» por cada Sub-eje.',
    pasos: [
      { clave: 'archivo', tipo: 'archivo', ext: XLS, titulo: 'Reporte diario GTR', ayuda: 'El mismo reporte diario que usas en «Sin Información».' },
      { clave: 'hora', tipo: 'hora', titulo: 'Hora del reporte', ayuda: 'Ej: 21:00' }] },
  { id: 'semanal', ic: '📆', nombre: 'Consolidador Semanal', corto: 'Une los Excel diarios de alertas en uno',
    necesitas: ['Los Excel diarios de Data Alertas (normalmente 6 o 7)'],
    obtienes: 'Un Excel semanal con reportes renumerados, ordenados por fecha y sin duplicados (y su CSV).',
    pasos: [{ clave: 'archivos', tipo: 'archivos', ext: XLS, titulo: 'Excel diarios', ayuda: 'Toca «Agregar» y selecciona varios a la vez, o agrégalos uno por uno.' }] },
  { id: 'mensual', ic: '🗓️', nombre: 'Consolidador Mensual', corto: 'Une los Excel semanales en uno',
    necesitas: ['Los Excel semanales generados con el Consolidador Semanal'],
    obtienes: 'Un Excel mensual consolidado (y su CSV).',
    pasos: [{ clave: 'archivos', tipo: 'archivos', ext: XLS, titulo: 'Excel semanales', ayuda: 'Selecciona todos los semanales del mes.' }] },
  { id: 'desconexion', ic: '🔌', nombre: 'Actualizar Desconexión', corto: 'PDF de desconexión → Excel actualizado',
    necesitas: ['PDF «Resumen de unidades que no han reportado»', 'Excel de desconexión existente (o crear uno nuevo)', 'Fecha actual', 'Opcional: Excel de Disponibilidad para el Status'],
    obtienes: 'El Excel de desconexión actualizado, listo para guardar o compartir.',
    pasos: [
      { clave: 'modo', tipo: 'opciones', soloUI: true, titulo: '¿Qué quieres hacer?', ayuda: 'Puedes actualizar el Excel que ya tienes o empezar uno nuevo.',
        opciones: [['actualizar', '📂 Actualizar mi Excel', 'Tengo el Excel de desconexión anterior'], ['nuevo', '🆕 Crear uno nuevo', 'Empezar desde cero']] },
      { clave: 'pdf', tipo: 'archivo', ext: ['pdf'], titulo: 'PDF de unidades sin reportar', ayuda: 'Es el PDF «Resumen de Unidades que No Han Reportado Hoy en la Última Hora» del GTRMax.' },
      { clave: 'excel', tipo: 'archivo', ext: XLS, si: d => d.modo === 'actualizar', titulo: 'Tu Excel de desconexión actual', ayuda: 'El archivo que quieres actualizar.' },
      { clave: 'responsable', tipo: 'texto', si: d => d.modo === 'nuevo', opcional: true, titulo: 'Responsable', ph: 'Nombre del responsable', ayuda: 'Aparecerá en las filas nuevas.' },
      { clave: 'corte', tipo: 'texto', si: d => d.modo === 'nuevo', opcional: true, titulo: 'Corte', ph: 'Nombre del corte', ayuda: 'Nombre del corte para las filas nuevas.' },
      { clave: 'fecha', tipo: 'fecha', titulo: 'Fecha actual', ayuda: 'Formato DD/MM/AAAA o DD/MM/AAAA HH:MM.' },
      { clave: 'status_excel', tipo: 'archivo', ext: XLS, opcional: true, titulo: 'Excel de Disponibilidad (opcional)', ayuda: 'Si lo adjuntas, se actualiza el Status de cada unidad. Si no, queda como estaba.' }] },
  { id: 'alertas', ic: '📝', nombre: 'Data Alertas', corto: 'Texto de WhatsApp → Excel',
    necesitas: ['Los reportes copiados desde WhatsApp (puedes pegar muchos juntos)', 'Opcional: el Excel anterior para añadirle los nuevos'],
    obtienes: 'Un Excel con una fila por reporte: responsable, eje, unidad, incidencia, nivel…',
    pasos: [
      { clave: 'texto', tipo: 'area', titulo: 'Pega los reportes', ph: 'Mantén pulsado aquí y elige «Pegar»…', ayuda: 'Copia los mensajes del grupo de WhatsApp y pégalos aquí. Cada uno debe traer su hora y fecha, por ejemplo [9:21 a. m., 11/8/2026].',
        ejemplo: '[9:21 a. m., 11/8/2026] Lon: Buenos Días\n* Eje Blv - Sub Eje Acevedo\n* Unidad 0645 - Vin 0684\n* CTS: Johan S.\nIncidencia: Exceso de velocidad\nIncumplimiento: 93 km/h en autopista\n[9:40 a. m., 11/8/2026] Lon: Buenas Tardes\n* Eje Pz\n* Unidad 0637 - Vin 0696\n* CTS: Wilfredo A.\nIncidencia: Exceso de velocidad\nIncumplimiento: 101 km/h en autopista' },
      { clave: 'excel', tipo: 'archivo', ext: XLS, opcional: true, titulo: 'Excel existente (opcional)', ayuda: 'Si lo adjuntas, los reportes nuevos se agregan a él. Si no, se crea uno nuevo.' }] },
  { id: 'ojo', ic: '🦅', nombre: 'Ojo de Halcón', corto: 'Cruza Estatus + Disponibilidad y detecta incidencias',
    necesitas: ['«Reporte de Estatus Actual» (.xls)', 'Excel de Disponibilidad', 'Opcional: unidades a obviar'],
    obtienes: 'Un mensaje por incidencia: fuera de ruta, movimiento no reportado, exceso de velocidad, sin movimiento o UT pegada.',
    pasos: [
      { clave: 'estatus', tipo: 'archivo', ext: XLS, titulo: 'Reporte de Estatus Actual', ayuda: 'El mismo que usas en «Últimas UT en Movimiento».' },
      { clave: 'disponibilidad', tipo: 'archivo', ext: XLS, titulo: 'Excel de Disponibilidad', ayuda: 'Excel con una hoja por Eje/Sub-eje (Acevedo, Brion, Buroz, Paez…).' },
      { clave: 'excluir', tipo: 'texto', opcional: true, titulo: '¿Obviar alguna UT? (opcional)', ph: 'Ej: R-036, 0140', ayuda: 'Separadas por coma.' }] },
  { id: 'recorrido', ic: '🧭', nombre: 'Resumen de Recorrido', corto: 'Recorrido y paradas de una unidad',
    necesitas: ['Historial de la unidad (CSV o Excel)', 'Número de unidad', 'Nombre del CTS'],
    obtienes: 'El resumen del recorrido, de las paradas o ambos, listo para enviar.',
    pasos: [
      { clave: 'archivo', tipo: 'archivo', ext: ['csv', 'xls', 'xlsx'], titulo: 'Historial de la unidad', ayuda: 'Exporta el historial de recorrido de UNA unidad.' },
      { clave: 'unidad', tipo: 'unidad', titulo: '¿De qué unidad es?', ph: 'Ej: R-066 o 0343', ayuda: 'Escribe el número como lo conoces: R-066, 0343…' },
      { clave: 'cts', tipo: 'texto', titulo: 'Nombre del CTS', ph: 'Ej: Johan Sosa', ayuda: 'Se mostrará como «Johan S.».' },
      { clave: 'tipo', tipo: 'opciones', titulo: '¿Qué reporte necesitas?', ayuda: 'Elige uno.',
        opciones: [['recorrido', '🧭 Resumen del recorrido', 'Por dónde pasó la unidad'], ['paradas', '⏱️ Resumen de paradas', 'Dónde y cuánto se detuvo'], ['ambos', '📄 Ambos', 'Recorrido y paradas']] }] }
];

/* ── Estado y navegación ───────────────────────────────────────────────────── */
const S = { pant: 'splash', t: null, paso: 0, d: {}, manualVisible: false, ocupado: false };
const vista = $('#vista');

function mostrar(nodo, dir = 'der') {
  nodo.classList.add('pantalla'); if (dir === 'izq') nodo.classList.add('izq');
  vista.replaceChildren(nodo); vista.scrollTop = 0;
  const conBarra = !['splash', 'onboarding'].includes(S.pant);
  $('#barra').classList.toggle('oculto', !conBarra);
  $('#btn-atras').style.visibility = S.pant === 'inicio' ? 'hidden' : 'visible';
  $('#btn-ayuda').style.visibility = S.pant === 'ayuda' ? 'hidden' : 'visible';
}
window.__atras = () => {
  if (S.ocupado) { toast('Procesando… espera un momento'); return true; }
  switch (S.pant) {
    case 'splash': case 'onboarding': case 'inicio': return false;
    case 'intro': case 'resultado': case 'error': return irInicio('izq'), true;
    case 'ayuda': return irInicio('izq'), true;
    case 'asistente': if (S.paso > 0) { S.paso = antes(S.paso, -1); pintarPaso('izq'); } else irIntro(S.t, 'izq'); return true;
  }
  return false;
};
$('#btn-atras').addEventListener('click', () => { if (!window.__atras()) window.Android.salir(); });
$('#btn-ayuda').addEventListener('click', () => irAyuda());

/* ── Splash + tutorial ─────────────────────────────────────────────────────── */
function irSplash() {
  S.pant = 'splash';
  mostrar(h('div', { class: 'splash' }, h('img', { src: 'img/logo.png', alt: 'ControlPlus CCO' }),
    h('div', { class: 'lema' }, 'Gestión y seguimiento de unidades'), h('div', { class: 'barra-carga' }, h('i'))));
  setTimeout(() => { let visto = false; try { visto = localStorage.getItem('cp_onb') === '1'; } catch (e) { } visto ? irInicio() : irOnboarding(0); }, 1900);
}
const ONB = [
  ['📱', 'Todo en tu teléfono', 'ControlPlus CCO funciona sin internet. Tus reportes se procesan aquí mismo y no salen del equipo.'],
  ['📂', 'Elige y sube tus reportes', 'Escoge una herramienta, adjunta el archivo del GTRMax y la app te guía paso a paso con ayudas claras.'],
  ['📤', 'Copia, comparte o guarda', 'Obtén mensajes listos para WhatsApp o archivos Excel que se guardan en Descargas/ControlPlus.']];
function irOnboarding(i) {
  S.pant = 'onboarding';
  const [e, t, p] = ONB[i], ultimo = i === ONB.length - 1;
  const fin = () => { try { localStorage.setItem('cp_onb', '1'); } catch (x) { } irInicio(); };
  mostrar(h('div', { class: 'onb' }, h('div', { class: 'emoji' }, e), h('h2', {}, t), h('p', {}, p),
    h('div', { class: 'puntos' }, ONB.map((_, k) => h('i', { class: k === i ? 'act' : '' }))),
    h('div', { class: 'fila-botones' },
      !ultimo && h('button', { class: 'btn sec', onclick: fin }, 'Omitir'),
      h('button', { class: 'btn naranja', onclick: () => ultimo ? fin() : irOnboarding(i + 1) }, ultimo ? '¡Empezar!' : 'Siguiente'))));
}

/* ── Inicio ────────────────────────────────────────────────────────────────── */
function irInicio(dir = 'der') {
  S.pant = 'inicio'; S.ocupado = false; S.t = null; S.d = {}; S.paso = 0;
  const cont = h('div', {},
    h('div', { class: 'hero' }, h('img', { src: 'img/emblema.png', alt: '' }),
      h('div', {}, h('h1', {}, 'ControlPlus ', h('b', {}, 'CCO')), h('p', {}, 'Gestión y seguimiento de unidades'))),
    h('div', { class: 'chips' },
      h('div', { class: 'chip' }, h('i', { class: 'punto' }), h('span', {}, 'Sin conexión · todo se procesa aquí')),
      h('div', { class: 'chip', id: 'chip-motor' }, h('i', { class: 'punto espera' }), h('span', {}, 'Preparando motor…'))),
    h('h3', { class: 'seccion' }, 'Herramientas'),
    h('div', { class: 'rejilla' }, T.map((t, i) => h('button', { class: 'tarjeta', style: `animation-delay:${80 + i * 55}ms`, onclick: () => irIntro(t) },
      h('div', { class: 'ic' }, t.ic), h('span', { class: 'num' }, i + 1), h('b', {}, t.nombre), h('span', { class: 'd' }, t.corto)))));
  mostrar(cont, dir); pintarMotor();
}

/* ── Introducción didáctica de cada herramienta ────────────────────────────── */
function irIntro(t, dir = 'der') {
  S.pant = 'intro'; S.t = t; S.d = {}; S.paso = 0; S.manualVisible = false;
  mostrar(h('div', {}, h('div', { class: 'card' },
    h('div', { class: 'etq' }, 'Herramienta ' + (T.indexOf(t) + 1)), h('h2', {}, t.ic + ' ' + t.nombre), h('p', { class: 'muted' }, t.corto),
    h('h3', { class: 'seccion', style: 'margin:14px 0 4px' }, 'Necesitarás'),
    h('ul', { class: 'lista-ico' }, t.necesitas.map(x => h('li', {}, '✔️', h('span', {}, x)))),
    h('h3', { class: 'seccion', style: 'margin:14px 0 4px' }, 'Obtendrás'), h('p', { style: 'margin:4px 0 0' }, t.obtienes)),
    h('button', { class: 'btn naranja', style: 'width:100%', onclick: () => { S.paso = antes(-1, 1); irAsistente(); } }, 'Empezar')), dir);
}

/* ── Asistente por pasos ───────────────────────────────────────────────────── */
const visibles = () => S.t.pasos.map((p, i) => [p, i]).filter(([p]) => !p.si || p.si(S.d)).map(([, i]) => i);
function antes(actual, delta) { const v = visibles(); const pos = v.indexOf(actual); return v[Math.max(0, Math.min(v.length - 1, (pos < 0 ? (delta > 0 ? -1 : v.length) : pos) + delta))]; }
function irAsistente() { S.pant = 'asistente'; pintarPaso('der'); }

function pintarPaso(dir = 'der', soloContenido = false) {
  const t = S.t, p = t.pasos[S.paso], v = visibles(), pos = v.indexOf(S.paso), total = v.length, ultimo = pos === total - 1;
  const error = h('div', { class: 'error-campo' });
  const cuerpo = h('div', { class: 'card', id: 'cuerpo-paso' },
    h('div', { class: 'etq' }, p.opcional ? 'Opcional' : 'Obligatorio'), h('h2', {}, p.titulo), campo(p, error),
    p.ayuda && h('div', { class: 'ayuda', html: '<b>💡 Ayuda:</b> ' + esc(p.ayuda) }), error);
  if (soloContenido && $('#cuerpo-paso')) { $('#cuerpo-paso').replaceWith(cuerpo); cuerpo.style.animation = 'none'; return; }
  const btnSig = h('button', { class: 'btn ' + (ultimo ? 'naranja' : ''), id: 'btn-sig' }, ultimo ? 'Generar' : 'Continuar');
  btnSig.addEventListener('click', () => avanzar(btnSig));
  const nodo = h('div', {}, h('div', { class: 'paso-txt' }, `${t.nombre} · Paso ${pos + 1} de ${total}`),
    h('div', { class: 'progreso' }, h('i', { style: `width:${((pos + 1) / total) * 100}%` })), cuerpo,
    h('div', { class: 'nav-inf' }, h('button', { class: 'btn sec', onclick: () => window.__atras() }, 'Atrás'), btnSig));
  mostrar(nodo, dir);
}

function campo(p, error) {
  const refrescar = () => pintarPaso('der', true);
  if (p.tipo === 'archivo') {
    const a = S.d[p.clave];
    return h('div', {}, !a && h('button', { class: 'selector', onclick: async () => {
        const r = await llamar('elegirArchivos', false); if (!r.ok || !r.archivos.length) { if (r.error) error.textContent = r.error; return; }
        if (!extOk(r.archivos[0].nombre, p.ext)) { error.textContent = 'Ese archivo no parece del tipo correcto (se espera: ' + p.ext.join(', ') + ').'; return; }
        S.d[p.clave] = r.archivos[0]; refrescar(); } },
      h('span', { class: 'ic' }, '📂'), 'Toca para elegir el archivo', h('small', { class: 'muted' }, 'Tipos: ' + p.ext.join(', '))),
      a && h('div', { class: 'archivo-ok' }, svgCheck(), h('div', { class: 'nm' }, a.nombre, h('small', {}, kb(a.bytes))),
        h('button', { 'aria-label': 'Quitar', onclick: () => { delete S.d[p.clave]; refrescar(); } }, '✕')));
  }
  if (p.tipo === 'archivos') {
    const lista = S.d[p.clave] || [];
    return h('div', {}, h('button', { class: 'selector', onclick: async () => {
        const r = await llamar('elegirArchivos', true); if (!r.ok) { error.textContent = r.error || ''; return; }
        const malos = r.archivos.filter(f => !extOk(f.nombre, p.ext)); const buenos = r.archivos.filter(f => extOk(f.nombre, p.ext));
        S.d[p.clave] = [...lista, ...buenos.filter(f => !lista.some(x => x.nombre === f.nombre))];
        refrescar(); if (malos.length) $('.error-campo').textContent = 'Se omitieron archivos que no son Excel: ' + malos.map(m => m.nombre).join(', '); } },
      h('span', { class: 'ic' }, '📚'), lista.length ? 'Agregar más archivos' : 'Toca para agregar archivos'),
      h('div', { class: 'lista-arch' }, lista.map((a, i) => h('div', { class: 'archivo-ok' }, svgCheck(), h('div', { class: 'nm' }, a.nombre, h('small', {}, kb(a.bytes))),
        h('button', { 'aria-label': 'Quitar', onclick: () => { lista.splice(i, 1); refrescar(); } }, '✕')))),
      lista.length > 0 && h('p', { class: 'muted', style: 'margin:10px 0 0;font-size:13px' }, lista.length + ' archivo(s) seleccionado(s)'));
  }
  if (p.tipo === 'opciones') {
    return h('div', { class: 'opciones' }, p.opciones.map(([val, tit, sub]) => h('button', { class: 'opcion' + (S.d[p.clave] === val ? ' sel' : ''), onclick: () => { S.d[p.clave] = val; refrescar(); } },
      h('div', {}, h('b', {}, tit), h('span', { class: 'muted', style: 'font-size:13px' }, sub)))));
  }
  const esArea = p.tipo === 'area';
  const val = S.d[p.clave] ?? (p.tipo === 'hora' ? horaAhora() : p.tipo === 'fecha' ? fechaAhora() : '');
  if (S.d[p.clave] === undefined && val) S.d[p.clave] = val;
  const ent = h(esArea ? 'textarea' : 'input', { class: 'txt', placeholder: p.ph || '', autocomplete: 'off', value: esArea ? false : val,
    inputmode: p.tipo === 'hora' ? 'text' : false, oninput: e => { S.d[p.clave] = e.target.value; error.textContent = ''; } });
  if (esArea) ent.value = val;
  const extra = [];
  if (p.tipo === 'unidad' && S.manualVisible) extra.push(
    h('div', { class: 'ayuda', style: 'margin-top:14px', html: '<b>Unidad nueva:</b> no está en mi lista de ejes/VIN. Escribe: <b>Eje, Sub eje, VIN</b> (ej. Blv, Acevedo, 1234; para Pz: Pz, 1234).' }),
    h('input', { class: 'txt', placeholder: 'Blv, Acevedo, 1234', value: S.d.manual || '', oninput: e => { S.d.manual = e.target.value; } }));
  return h('div', {}, ent, esArea && p.ejemplo && h('button', { class: 'btn sec chico', style: 'margin-top:10px', onclick: () => { S.d[p.clave] = p.ejemplo; ent.value = p.ejemplo; } }, 'Cargar texto de ejemplo'), extra);
}
const extOk = (n, ex) => !ex || ex.includes(n.split('.').pop().toLowerCase());
const svgCheck = () => { const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('class', 'check'); s.innerHTML = '<path d="M4 12.5l5 5L20 6.5"/>'; return s; };

async function avanzar(btn) {
  const t = S.t, p = t.pasos[S.paso], err = $('.error-campo'), v = visibles(), ultimo = v.indexOf(S.paso) === v.length - 1;
  const dato = S.d[p.clave];
  if (!p.opcional) {
    if ((p.tipo === 'archivo' && !dato) || (p.tipo === 'archivos' && !(dato && dato.length))) { err.textContent = p.tipo === 'archivos' ? 'Agrega al menos un archivo para continuar.' : 'Elige el archivo para continuar.'; return shake(); }
    if (['hora', 'fecha', 'texto', 'area', 'unidad'].includes(p.tipo) && !String(dato || '').trim()) { err.textContent = 'Este dato es obligatorio.'; return shake(); }
    if (p.tipo === 'opciones' && !dato) { err.textContent = 'Elige una opción para continuar.'; return shake(); }
  }
  if (p.tipo === 'unidad') {
    btn.classList.add('cargando');
    const r = await llamar('unidadConocida', S.d[p.clave]); btn.classList.remove('cargando');
    if (!r.alias) { err.textContent = 'No reconozco esa unidad. Escríbela como R-066 o 0343.'; return shake(); }
    if (!r.info && !String(S.d.manual || '').trim()) { S.manualVisible = true; pintarPaso('der', true); $('.error-campo').textContent = 'Esa unidad no está en la lista: completa los datos manuales.'; return; }
  }
  if (ultimo) return generar();
  S.paso = antes(S.paso, 1); pintarPaso('der');
}
function shake() { const c = $('#cuerpo-paso'); c && c.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-8px)' }, { transform: 'translateX(8px)' }, { transform: 'translateX(0)' }], { duration: 280 }); }

function armarParams() {
  const p = {};
  for (const s of S.t.pasos) {
    if (s.soloUI) continue; if (s.si && !s.si(S.d)) continue;
    const v = S.d[s.clave]; if (v == null || v === '' || (Array.isArray(v) && !v.length)) continue;
    p[s.clave] = s.tipo === 'archivo' ? v.ruta : s.tipo === 'archivos' ? v.map(x => x.ruta) : v;
  }
  if (S.t.id === 'recorrido' && S.d.manual) p.manual = S.d.manual;
  return p;
}

/* ── Procesando ────────────────────────────────────────────────────────────── */
const TIPS = ['Leyendo tu archivo…', 'Aplicando las reglas de negocio…', 'Armando los mensajes…', 'Verificando los datos…', 'Casi listo…'];
async function generar() {
  S.pant = 'cargando'; S.ocupado = true;
  const tip = h('div', { class: 'tip' }, TIPS[0]); let i = 0;
  const timer = setInterval(() => { i = (i + 1) % TIPS.length; tip.textContent = TIPS[i]; tip.style.animation = 'none'; tip.offsetWidth; tip.style.animation = ''; }, 2400);
  mostrar(h('div', { class: 'cargando-pant' }, h('div', { class: 'orbita' }, h('div', { class: 'anillo' }), h('img', { src: 'img/emblema.png', alt: '' })), tip,
    h('div', { class: 'nota' }, 'La primera vez puede tardar un poco más mientras se prepara el motor. No cierres la app.')));
  const r = await llamar('ejecutar', S.t.id, JSON.stringify(armarParams()));
  clearInterval(timer); S.ocupado = false; r.ok ? irResultado(r) : irError(r);
}

/* ── Resultado / error ─────────────────────────────────────────────────────── */
function formatoWA(txt) { // negrita/cursiva como WhatsApp: *x* y _x_ solo si pegados a texto
  return esc(txt).replace(/(^|[\s(])\*(?=\S)([^*\n]*?\S)\*(?=$|[\s.,;:!?)])/g, '$1<b>$2</b>').replace(/(^|[\s(])_(?=\S)([^_\n]*?\S)_(?=$|[\s.,;:!?)])/g, '$1<i>$2</i>');
}
const checkGrande = () => { const s = svgCheck(); return s; };
function irResultado(r) {
  S.pant = 'resultado';
  const msgs = r.mensajes || [], arch = r.archivos || [];
  const nodo = h('div', {}, PREVIEW && h('div', { class: 'banner-demo' }, '👀 Vista previa en navegador: resultados de demostración, no del motor real.'));
  if (r.vacio) {
    nodo.append(h('div', { class: 'card' }, h('div', { class: 'resumen vacio' }, h('div', { class: 'ok-ic' }, '✨'), h('div', {}, h('h2', { style: 'font-size:18px' }, 'Todo en orden'), h('div', { class: 'muted' }, r.vacio)))));
  } else {
    const n = msgs.length || arch.length;
    nodo.append(h('div', { class: 'card' }, h('div', { class: 'resumen' }, h('div', { class: 'ok-ic' }, checkGrande()),
      h('div', {}, h('h2', { style: 'font-size:19px' }, msgs.length ? `${n} mensaje(s) listo(s)` : 'Archivo generado'),
        h('div', { class: 'muted' }, r.total_registros ? `${r.total_registros} registro(s) procesado(s)` : S.t.nombre))),
      msgs.length > 1 && h('div', { class: 'fila-botones', style: 'margin-top:12px' },
        h('button', { class: 'btn sec chico', onclick: () => { window.Android.copiar(msgs.map(m => m.texto).join('\n\n')); toast('Todos los mensajes copiados'); } }, '📋 Copiar todos'))));
  }
  msgs.forEach((m, i) => nodo.append(h('div', { class: 'msg', style: `animation-delay:${i * 60}ms` },
    h('div', { class: 'msg-cab' }, m.titulo || 'Mensaje ' + (i + 1)),
    h('pre', { html: formatoWA(m.texto) }),
    h('div', { class: 'msg-acc' },
      h('button', { class: 'btn chico', onclick: () => { window.Android.copiar(m.texto); toast('Copiado ✓'); } }, '📋 Copiar'),
      h('button', { class: 'btn chico naranja', onclick: () => window.Android.compartirTexto(m.texto) }, '📤 Compartir')))));
  arch.forEach((ruta, i) => { const nombre = ruta.split(/[\\/]/).pop(); nodo.append(h('div', { class: 'msg', style: `animation-delay:${i * 60}ms` },
    h('div', { class: 'archivo-res' }, h('div', { class: 'ic' }, nombre.endsWith('.csv') ? '📄' : '📊'), h('div', { class: 'nm' }, nombre)),
    h('div', { class: 'msg-acc' },
      h('button', { class: 'btn chico', onclick: () => { const g = JSON.parse(window.Android.guardarArchivo(ruta)); toast(g.ok ? 'Guardado en ' + g.carpeta : 'No se pudo guardar: ' + g.error); } }, '💾 Guardar'),
      h('button', { class: 'btn chico naranja', onclick: () => window.Android.compartirArchivo(ruta) }, '📤 Compartir')))); });
  if (r.logs && r.logs.length) nodo.append(h('details', { class: 'log' }, h('summary', {}, 'Ver detalles del proceso'), h('pre', {}, r.logs.join('\n'))));
  nodo.append(h('div', { class: 'nav-inf' }, h('button', { class: 'btn sec', onclick: () => irIntro(S.t) }, '🔁 Repetir'), h('button', { class: 'btn', onclick: () => irInicio('izq') }, 'Inicio')));
  mostrar(nodo);
}
function irError(r) {
  S.pant = 'error';
  mostrar(h('div', {}, h('div', { class: 'error-caja' }, h('h2', {}, '⚠️ No se pudo completar'), h('p', { style: 'margin:0 0 8px' }, r.error || 'Ocurrió un problema inesperado.'),
    h('div', { class: 'muted', style: 'font-size:13px' }, 'Revisa que sean los archivos correctos y vuelve a intentar. Si el problema continúa, usa «Ver detalles».')),
    r.detalle && h('details', { class: 'log' }, h('summary', {}, 'Ver detalles técnicos'), h('pre', {}, r.detalle)),
    h('div', { class: 'nav-inf' }, h('button', { class: 'btn sec', onclick: () => irInicio('izq') }, 'Inicio'),
      h('button', { class: 'btn naranja', onclick: () => { S.paso = antes(-1, 1); irAsistente(); } }, 'Reintentar'))));
}

/* ── Acerca de / ayuda ─────────────────────────────────────────────────────── */
function irAyuda() {
  if (S.ocupado) return; const previa = S.pant; S.pant = 'ayuda';
  const c = motor.info && motor.info.compat ? Object.entries(motor.info.compat).map(([k, v]) => `${k}: ${v}`).join(' · ') : (motor.info && motor.info.error) || 'iniciando…';
  mostrar(h('div', {}, h('div', { class: 'card' }, h('img', { src: 'img/logo.png', alt: '', style: 'width:70%;max-width:260px;display:block;margin:0 auto 8px' }),
    h('h2', { style: 'text-align:center' }, 'ControlPlus CCO'), h('p', { class: 'muted', style: 'text-align:center;margin:0' }, 'Versión ' + VERSION)),
    h('div', { class: 'card' }, h('div', { class: 'etq' }, 'Privacidad'), h('ul', { class: 'lista-ico' },
      h('li', {}, '📴', h('span', {}, 'La app no tiene permiso de internet: nada sale de tu teléfono.')),
      h('li', {}, '🧹', h('span', {}, 'Los archivos que eliges se copian a una carpeta temporal que se borra al abrir la app de nuevo.')),
      h('li', {}, '💾', h('span', {}, 'Solo se guarda en Descargas/ControlPlus lo que tú decides guardar.'))),
      h('div', { class: 'etq', style: 'margin-top:14px' }, 'Motor'), h('p', { class: 'muted', style: 'margin:0;font-size:13px;word-break:break-word' }, c)),
    h('div', { class: 'fila-botones' }, h('button', { class: 'btn sec', onclick: () => irOnboarding(0) }, '🎓 Ver tutorial'), h('button', { class: 'btn', onclick: () => irInicio('izq') }, 'Volver')),
    h('p', { class: 'muted', style: 'text-align:center;font-size:12px;margin-top:18px' }, 'Desarrollado por MoussaCorp · Se reservan todos los derechos')));
}

irSplash();
