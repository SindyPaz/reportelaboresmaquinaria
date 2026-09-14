// ══════════════════════════════════════════════════════════════
// CREAR HOJAS DE LISTAS PARA APP1 — Anomalías Maquinaria
// Ejecutar UNA SOLA VEZ desde Apps Script
// Nota: Lista_operadores, Lista_equipos, Lista_implementos,
//       Lista_lotes y Areas ya existen desde el setup de APP2.
// ══════════════════════════════════════════════════════════════

function crearHojasApp1() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── 1. Hoja Fallas ──────────────────────────────────────────
  let hf = ss.getSheetByName('Fallas') || ss.insertSheet('Fallas');
  hf.clearContents();
  const encFallas = [['SISTEMA','SUBSISTEMA','PARTE','MODO DE FALLA','MECANISMO DE FALLA']];
  hf.getRange(1,1,1,5).setValues(encFallas);
  hf.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
  // Ejemplo de filas — reemplazar o completar con tus datos reales
  const filasEjemplo = [
    ['Chasis-estructura','Subsistema 3: Sistema de enganche','Abrazadera de boom','desajuste','Vibraciones/fatiga/desajuste'],
    ['Hidraulico','Subsistema 2: Depósito y gestión del fluido','Aceite hidráulico','fuga','Desgaste interno/carga variable'],
    ['Transmision','Subsistema 4: Toma de fuerza (PTO)','Acople hidráulico de PTO','fuga','Junta desgastada/deteriorada'],
  ];
  hf.getRange(2,1,filasEjemplo.length,5).setValues(filasEjemplo);
  Logger.log('✓ Hoja Fallas creada');

  // ── 2. Hoja Técnico ─────────────────────────────────────────
  let ht = ss.getSheetByName('Técnico') || ss.insertSheet('Técnico');
  ht.clearContents();
  ht.getRange(1,1).setValue('TECNICO');
  ht.getRange(1,1).setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
  const tecnicos = [
    ['Yeison Alvarez'],['Esteban Correa'],['Cesar Manco'],['Marco Parada'],
    ['David Urrego'],['Jhon Alvarez'],['Camilo Salcedo'],['Freddy Muñoz'],
    ['Rodrigo García'],['Jhon Quintero'],['Alfredo Castellar'],['Arley Serrano'],
    ['Carlos Mosquera'],['Daniel Ruiz'],['Diego Cardona'],['Edwin Torres'],
    ['Fernando Lopez'],['Hernán Castillo'],['Ivan Moreno'],['Jorge Perez'],['Luis Ramirez']
  ];
  ht.getRange(2,1,tecnicos.length,1).setValues(tecnicos);
  Logger.log('✓ Hoja Técnico creada');

  // ── 3. Hoja Prioridad ────────────────────────────────────────
  let hp = ss.getSheetByName('Prioridad') || ss.insertSheet('Prioridad');
  hp.clearContents();
  hp.getRange(1,1).setValue('NIVEL DE PRIORIDAD');
  hp.getRange(1,1).setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
  hp.getRange(2,1,3,1).setValues([['Alta'],['Media'],['Baja']]);
  Logger.log('✓ Hoja Prioridad creada');

  // ── 4. Hoja Disponibilidad ───────────────────────────────────
  let hd = ss.getSheetByName('Disponibilidad') || ss.insertSheet('Disponibilidad');
  hd.clearContents();
  hd.getRange(1,1).setValue('DISPONIBILIDAD');
  hd.getRange(1,1).setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
  hd.getRange(2,1,2,1).setValues([['Varada'],['Disponible']]);
  Logger.log('✓ Hoja Disponibilidad creada');

  SpreadsheetApp.getUi().alert(
    '✓ Hojas APP1 creadas:\n' +
    '• Fallas (completa con tus datos reales)\n' +
    '• Técnico\n• Prioridad\n• Disponibilidad\n\n' +
    'Recuerda llenar la hoja Fallas con todas las filas del LISTAS.xlsx'
  );
}
