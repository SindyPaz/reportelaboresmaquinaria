// Ejecutar una sola vez para agregar hojas Unidad y Areas
function agregarHojasUnidadAreas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Hoja Unidad
  let hu = ss.getSheetByName('Unidad') || ss.insertSheet('Unidad');
  hu.clearContents();
  hu.getRange(1,1).setValue('UNIDAD');
  hu.getRange(1,1).setFontWeight('bold').setBackground('#065F46').setFontColor('#FFFFFF');
  const unidades = [['UNIDAD'], ['METROS'], ['HAS'], ['BULTOS'], ['UNIDADES'], ['LITROS'], ['BINES'], ['VIAJES'], ['KILOS'], ['TONELADAS']];
  hu.getRange(2, 1, unidades.length, 1).setValues(unidades);
  
  // Hoja Areas
  let ha = ss.getSheetByName('Lista_areas') || ss.insertSheet('Lista_areas');
  ha.clearContents();
  ha.getRange(1,1).setValue('AREA');
  ha.getRange(1,1).setFontWeight('bold').setBackground('#065F46').setFontColor('#FFFFFF');
  const areas = [['Aplicaciones'], ['CTIA'], ['Control operaciones'], ['Cosecha'], ['Desarrollo de cultivo'], ['Empaque'], ['Preparación'], ['Taller'], ['semillero'], ['siembra']];
  ha.getRange(2, 1, areas.length, 1).setValues(areas);
  
  Logger.log('✓ Hojas Unidad y Lista_areas creadas');
  SpreadsheetApp.getUi().alert('✓ Hojas Unidad y Lista_areas creadas correctamente');
}
