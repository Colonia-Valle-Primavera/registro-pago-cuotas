function execute_registry_addition() {
  // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  add_message_to_log("Procesando...");
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Adición de Registro", '¿Te gustaría añadir el registro llenado?', ui.ButtonSet.YES_NO,);

  // Process the user's response.
  if (response == ui.Button.YES) {
    add_registry();
  } else {
    add_message_to_log("El usuario decidió no continuar con la adición del registro.");
  }
}

function execute_registry_cleaning() {
  // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  add_message_to_log("Procesando...");
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Limpieza de Registro", '¿Te gustaría limpiar el registro actual?', ui.ButtonSet.YES_NO,);

  // Process the user's response.
  if (response == ui.Button.YES) {
    clean_registry();
  } else {
    add_message_to_log("El usuario decidió no continuar con la limpieza del registro.");
  }
}

function add_message_to_log(message) {
  var sheetRegistry = load_sheet("registro");
  Logger.log(message);
  sheetRegistry.getRange(15, 9).setValue(message);
}

function add_registry() {

  var sheetRegistry = load_sheet("registro");
  var sheetTabla = load_sheet("tabla");

  var folio = sheetRegistry.getRange(6, 4).getValue();
  var fecha = sheetRegistry.getRange(6, 7).getValue();
  var monto = sheetRegistry.getRange(6, 11).getValue();
  var aportador = sheetRegistry.getRange(8, 4).getValue();
  var periodo_mes_inicio = sheetRegistry.getRange(10, 7).getValue();
  var periodo_ano_inicio = sheetRegistry.getRange(10, 8).getValue();
  var periodo_mes_fin = sheetRegistry.getRange(10, 10).getValue();
  var periodo_ano_fin = sheetRegistry.getRange(10, 11).getValue();
  var recaudador1 = sheetRegistry.getRange(12, 4).getValue();
  var recaudador2 = sheetRegistry.getRange(12, 7).getValue();
  var metodo = sheetRegistry.getRange(12, 11).getValue();
  var created_at = new Date();
  var email = Session.getEffectiveUser().getEmail();

  validate_registry_variables(folio, fecha, monto, aportador, periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin, recaudador1, recaudador2, metodo, created_at, email);

  check_folio_is_not_duplicated(sheetTabla, folio);
  check_dates_make_sense(periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin);

  const zip = (...args) => args[0].map((_, i) => args.map(arg => arg[i]));

  var start_date = get_date_string(periodo_mes_inicio, periodo_ano_inicio);
  var end_date = get_date_string(periodo_mes_fin, periodo_ano_fin);
  var periodo_range = dateRange(start_date, end_date);

  var montos_range = get_array_of_amounts(monto, periodo_range.length);

  for ( const [monto_por_mes, fecha_por_mes] of zip(montos_range, periodo_range) ) {
    add_registry_line(sheetTabla, folio, fecha, monto, aportador, periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin, recaudador1, recaudador2, metodo, created_at, email, monto_por_mes, fecha_por_mes);
  }

  var successful_message = "El registro con el folio '" + folio + "' fue añadido con éxito.";
  add_message_to_log(successful_message);

}

function load_sheet(sheet_name) {

  switch (sheet_name) {
    case "registro":
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("registro");
      break;
    case "tabla":
      var sheet = SpreadsheetApp.openById("19rfJyEZBKAVmHUD8FfuZtVB1EW3-gCkmPAWvgivNTcE").getSheetByName("Tabla única (conglomerado)");
      break;
  }

  if (sheet != null) {
    return sheet;
  }

  var error_message = "No se puede cargar la hoja de: " + sheet_name;
  show_message_error(error_message);
}

function check_is_null(variable) {
  return ( variable == "" || variable === undefined || variable === null );
}

function check_is_nan(variable) {
  return ( check_is_null(variable) || isNaN(variable) );
}

function show_message_error(message) {
  message = "Error. " + message;
  add_message_to_log(message);
  throw message;
}

function validate_registry_variables(folio, fecha, monto, aportador, periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin, recaudador1, recaudador2, metodo, created_at, email) {

  if ( check_is_nan(folio) ) {
    show_message_error("El 'Folio' viene vacío.");
  }

  if ( check_is_nan(fecha) ) {
    show_message_error("La 'Fecha' viene vacía.");
  }

  if ( check_is_nan(monto) ) {
    show_message_error("El 'monto' viene vacío.");
  }

  if ( check_is_null(aportador) ) {
    show_message_error("El 'Aportador' viene vacío.");
  }

  if ( check_is_null(periodo_mes_inicio) ) {
    show_message_error("El 'periodo_mes_inicio' viene vacío.");
  }

  if ( check_is_nan(periodo_ano_inicio) ) {
    show_message_error("El 'periodo_ano_inicio' viene vacío.");
  }

  if ( check_is_null(periodo_mes_fin) ) {
    show_message_error("El 'periodo_mes_fin' viene vacío.");
  }

  if ( check_is_nan(periodo_ano_fin) ) {
    show_message_error("El 'periodo_ano_fin' viene vacío.");
  }

  if ( check_is_null(recaudador1) ) {
    show_message_error("El 'Recaudador1' viene vacío.");
  }

  if ( check_is_null(recaudador2) ) {
    show_message_error("El 'Recaudador2' viene vacío.");
  }

  if ( check_is_null(metodo) ) {
    show_message_error("El 'metodo' viene vacío.");
  }

  if ( check_is_nan(created_at) ) {
    show_message_error("No se pudo cargar la fecha de registro.");
  }

  if ( check_is_null(email) ) {
    show_message_error("No se pudo cargar el correo del usuario.");
  }

}

function check_dates_make_sense(periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin) {

  if ( periodo_ano_inicio > periodo_ano_fin ) {
    show_message_error("El año de inicio es mayor al año de fin.");
  }

  var mes_inicio = parseInt(get_numeric_month(periodo_mes_inicio));
  var mes_fin = parseInt(get_numeric_month(periodo_mes_fin));

  if ( periodo_ano_inicio == periodo_ano_fin && mes_inicio > mes_fin ) {
    show_message_error("El mes de inicio es mayor al mes de fin.");
  }
}

function is_true(value) {
  return value === true;
}

function check_folio_is_not_duplicated(sheetTabla, folio) {
  var lastRowIndex = sheetTabla.getLastRow();
  var folio_range_values = sheetTabla.getRange(2, 3, lastRowIndex-1).getValues();
  var is_folio_in_values = folio_range_values.map((x) => x[0] == folio);

  if ( is_folio_in_values.some(is_true) ) {
    var error_message = "El folio '" + folio + "' está duplicado. Verifique que el registro que quiere añadir es correcto."
    show_message_error(error_message);
  }
}

function get_formatted_date(dateToBeFormatted) {
  return Utilities.formatDate(dateToBeFormatted, 'Etc/UTC', 'yyyy-MM-dd');
}

function get_date_in_CST(dateToBeFormatted) {
  return Utilities.formatDate(dateToBeFormatted, 'Mexico/General', 'yyyy-MM-dd HH:mm:ss Z');
}

function get_numeric_month(month) {
  switch (month) {
    case "Enero":
      return "01";
    case "Febrero":
      return "02";
    case "Marzo":
      return "03";
    case "Abril":
      return "04";
    case "Mayo":
      return "05";
    case "Junio":
      return "06";
    case "Julio":
      return "07";
    case "Agosto":
      return "08";
    case "Septiembre":
      return "09";
    case "Octubre":
      return "10";
    case "Noviembre":
      return "11";
    case "Diciembre":
      return "12";
  }
}

function get_month_in_english(month) {
  switch (month) {
    case "Enero":
      return "Jan";
    case "Febrero":
      return "Feb";
    case "Marzo":
      return "Mar";
    case "Abril":
      return "Apr";
    case "Mayo":
      return "May";
    case "Junio":
      return "Jun";
    case "Julio":
      return "Jul";
    case "Agosto":
      return "Aug";
    case "Septiembre":
      return "Sep";
    case "Octubre":
      return "Oct";
    case "Noviembre":
      return "Nov";
    case "Diciembre":
      return "Dec";
  }
}

function get_date_string(mes, ano) {
  var date_string = String(ano) + "-" + get_numeric_month(mes) + "-01";
  return date_string
}

function dateRange(startDate, endDate) {
  // we use UTC methods so that timezone isn't considered
  let start = new Date(startDate);
  const end = new Date(endDate).setUTCHours(12);
  const dates = [];
  while (start <= end) {
    // compensate for zero-based months in display
    const displayMonth = start.getUTCMonth() + 1;
    dates.push([
      start.getUTCFullYear(),
      // months are zero based, ensure leading zero
      (displayMonth).toString().padStart(2, '0'),
      // always display the first of the month
      '01',
    ].join('-'));

    // progress the start date by one month
    start = new Date(start.setUTCMonth(displayMonth));
  }

  return dates;
}

function monthDiff(d1, d2) {
  // Wasn't used
  var months;
  months = (d2.getFullYear() - d1.getFullYear()) * 12;
  months -= d1.getMonth();
  months += d2.getMonth();
  return months <= 0 ? 0 : months;
}

function get_array_of_amounts(monto, num_of_months) {
  var amounts_array = [];

  if ( num_of_months == 12 && monto == 1100 ) {
    for ( var i=0; i<11; i++ ) {
      amounts_array.push(100);
    }
    amounts_array.push(0);
  } else {
    for (i=0; i<num_of_months; i++) {
      amounts_array.push(monto/num_of_months);
    }
  }
  return amounts_array;
}

function get_folio_integrado(folio, fecha_por_mes) {
  var folio_integrado = "I-" + folio + "-" + fecha_por_mes.slice(0,4) + fecha_por_mes.slice(5,7);
  return folio_integrado;
}

function get_meses_completos(periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin) {

  var meses_completos;

  if ( periodo_mes_inicio == periodo_mes_fin && periodo_ano_inicio == periodo_ano_fin ) {

    meses_completos = get_month_in_english(periodo_mes_inicio) + " " + periodo_ano_inicio;

  } else if ( periodo_ano_inicio == periodo_ano_fin ) {

    meses_completos = get_month_in_english(periodo_mes_inicio) + " - " + get_month_in_english(periodo_mes_fin) + " " + periodo_ano_fin;

  } else {

    meses_completos = get_month_in_english(periodo_mes_inicio) + " " + periodo_ano_inicio + " - " + get_month_in_english(periodo_mes_fin) + " " + periodo_ano_fin;

  }

  return meses_completos;
}

function add_registry_line(sheetTabla, folio, fecha, monto, aportador, periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin, recaudador1, recaudador2, metodo, created_at, email, monto_por_mes, fecha_por_mes) {

  var lastRowIndex = sheetTabla.getLastRow();

  sheetTabla.getRange(lastRowIndex + 1, 1).setValue(get_formatted_date(fecha));

  var folio_integrado = get_folio_integrado(folio, fecha_por_mes);

  sheetTabla.getRange(lastRowIndex + 1, 2).setValue(folio_integrado);
  sheetTabla.getRange(lastRowIndex + 1, 3).setValue(folio);
  sheetTabla.getRange(lastRowIndex + 1, 4).setValue("Ingreso");
  sheetTabla.getRange(lastRowIndex + 1, 6).setValue("I - Cuota");
  sheetTabla.getRange(lastRowIndex + 1, 7).setValue("Vigente");
  sheetTabla.getRange(lastRowIndex + 1, 8).setValue(metodo);
  sheetTabla.getRange(lastRowIndex + 1, 9).setValue(monto);
  sheetTabla.getRange(lastRowIndex + 1, 10).setValue(monto_por_mes);
  sheetTabla.getRange(lastRowIndex + 1, 11).setValue(aportador);
  
  var meses_completos = get_meses_completos(periodo_mes_inicio, periodo_ano_inicio, periodo_mes_fin, periodo_ano_fin);

  sheetTabla.getRange(lastRowIndex + 1, 12).setValue(meses_completos);
  sheetTabla.getRange(lastRowIndex + 1, 13).setValue(fecha_por_mes);
  sheetTabla.getRange(lastRowIndex + 1, 17).setValue(recaudador1);
  sheetTabla.getRange(lastRowIndex + 1, 18).setValue(recaudador2);
  sheetTabla.getRange(lastRowIndex + 1, 22).setValue(get_date_in_CST(created_at));
  sheetTabla.getRange(lastRowIndex + 1, 23).setValue(email);

}

function clean_registry() {

  var sheetRegistry = load_sheet("registro");
  sheetRegistry.getRange(6, 4).setValue(null);
  sheetRegistry.getRange(6, 7).setValue(null);
  sheetRegistry.getRange(6, 11).setValue(null);
  sheetRegistry.getRange(8, 4).setValue(null);
  sheetRegistry.getRange(10, 7).setValue(null);
  sheetRegistry.getRange(10, 8).setValue(null);
  sheetRegistry.getRange(10, 10).setValue(null);
  sheetRegistry.getRange(10, 11).setValue(null);
  sheetRegistry.getRange(12, 4).setValue(null);
  sheetRegistry.getRange(12, 7).setValue(null);
  sheetRegistry.getRange(12, 11).setValue(null);
  
  var cleaned_message = "Los campos del registro han sido limpiados con éxito.";
  add_message_to_log(cleaned_message);
}
