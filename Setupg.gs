// =================================================================
// SETUP FUNCTIONS — Ejecutar UNA sola vez cada una
// =================================================================

// ---- MASTER: Corre esto y se hace todo ----
function setupTemplateEngine() {
  Logger.log('🚀 === SETUP TEMPLATE ENGINE ===');
  
  // Paso 1: Crear template Cashback
  const cashbackDocId = seedCashbackTemplate();
  Logger.log('✅ Template Cashback creado: ' + cashbackDocId);
  
  // Paso 2: Crear template Concurso
  const concursoDocId = seedConcursoTemplate();
  Logger.log('✅ Template Concurso creado: ' + concursoDocId);
  
  // Paso 3: Registrar en Template_Registry
  _seedRegistry(cashbackDocId, concursoDocId);
  Logger.log('✅ Registry creado');
  
  // Paso 4: Crear campos en Template_Fields
  seedColombiaFields();
  Logger.log('✅ Campos creados');
  
  Logger.log('🎉 === SETUP COMPLETO ===');
  Logger.log('📋 Cashback Doc ID: ' + cashbackDocId);
  Logger.log('📋 Concurso Doc ID: ' + concursoDocId);
  Logger.log('');
  Logger.log('👉 SIGUIENTE PASO: Cambia 2 líneas en processWebPayload');
  Logger.log('   ANTES:  const v67Map = mapWebToEngine(payload);');
  Logger.log('           const result = coreEngine(v67Map, payload.userEmail);');
  Logger.log('   DESPUÉS: const result = coreEngineV2(payload, payload.userEmail);');
  
  return { cashbackDocId, concursoDocId };
}

// ---- CASHBACK TEMPLATE ----
function seedCashbackTemplate() {
  const doc = DocumentApp.create('Template_CO_Cashback');
  const body = doc.getBody();
  
  // TÍTULO
  const titulo = body.appendParagraph('TÉRMINOS Y CONDICIONES – CAMPAÑA "{{NOMBRE_CAMPANA_UPPER}}"');
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  
  // INTRODUCCIÓN
  appendRichParagraph(body, [
    'Por medio del presente documento se dan a conocer los términos y condiciones de la campaña denominada "',
    {text: '{{NOMBRE_CAMPANA_LOWER}}', bold: true},
    '" (en adelante la "Campaña"). La participación en la Campaña constituye la aceptación total e incondicional de los presentes Términos y Condiciones, los cuales resultan definitivos y vinculantes para los Usuarios/Consumidores participantes, que cumplan con los requisitos aquí dispuestos.'
  ]);
  
  // I. TERRITORIO
  appendRichSection(body, 'I. Territorio: ', [
    'La Campaña será válida únicamente para las órdenes realizadas dentro de las zonas de cobertura de {{REF_TIENDA}} en la Plataforma Rappi, en ',
    {text: '{{TEXTO_TERRITORIO}}', bold: true}, '.'
  ]);
  
  // II. VIGENCIA
  appendRichSection(body, 'II. Vigencia: ', [
    'Campaña válida desde las ', {text: '{{HORA_INICIO}}', bold: true},
    ' del ', {text: '{{FECHA_INICIO}}', bold: true},
    ' hasta las ', {text: '{{HORA_FIN}}', bold: true},
    ' del ', {text: '{{FECHA_FIN}}', bold: true},
    ' y/o hasta agotar existencias, lo que primero ocurra. Para efectos de la Campaña, se ha establecido un valor total máximo de ',
    {text: '{{PRESUPUESTO_LETRAS}}', bold: true}, ' (', {text: '{{PRESUPUESTO_NUM}}', bold: true},
    ') de Créditos a ser entregados como Cashback a los Usuarios/Consumidores que participen en la Campaña ("Existencias"). Por consiguiente, el agotamiento de las Existencias con anterioridad al final de la Vigencia indicada implicará la terminación de la Campaña.'
  ]);
  
  // III. TIPO DE USUARIOS
  appendRichSection(body, 'III. Tipo de Usuarios Participantes: ', ['{{TEXTO_SEGMENTO}}']);
  
  // IV. TIENDA PARTICIPANTE
  appendRichSection(body, '{{TITULO_TIENDA}}', [
    'Participarán todas las tiendas virtuales de ', {text: '{{TIENDA_DISPLAY}}', bold: true},
    '{{DEFINICION_TIENDA}}al interior de la Plataforma Rappi ubicadas dentro del Territorio.'
  ]);
  
  // V. PRODUCTOS PARTICIPANTES
  appendRichSection(body, 'V. Productos Participantes: ', [
    'Participarán todos los productos que hacen parte del catálogo de {{REF_TIENDA}} al interior de la Plataforma Rappi.'
  ]);
  
  // VI. BENEFICIO
  const pBeneficio = body.appendParagraph('');
  pBeneficio.appendText('VI. Beneficio: ').setBold(true);
  appendRichTextToParagraph(pBeneficio, [
    'Los Usuarios/Consumidores Participantes que durante la Vigencia de la Campaña compren cualquiera de los Productos Participantes de {{REF_TIENDA}} recibirán en Créditos el ',
    {text: '{{TEXTO_PORCENTAJE}}', bold: true},
    ' del valor de dichos Productos Participantes (en adelante el "Cashback"). Dichos Créditos serán cargados a su cuenta al interior de la Plataforma Rappi. Se aclara que el monto máximo del Cashback que se otorgará en Créditos es de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true}, ')',
    ' Créditos. Por lo tanto, en caso de que el Usuario/Consumidor realice una compra en {{REF_TIENDA}} por un valor superior a los ',
    {text: '{{UMBRAL_LETRAS}}', bold: true}, ' pesos M/Cte ($', {text: '{{UMBRAL_NUM}}', bold: true},
    '), recibirá un monto máximo en Créditos de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true}, ').',
    ' Los Créditos serán cargados a la cuenta de los Usuarios/Consumidores Participantes ',
    {text: '{{TEXTO_CARGA}}', bold: true},
    '. Los Créditos {{TEXTO_VIGENCIA_CREDITOS}}, entendiéndose que si el Usuario/Consumidor no hace uso de ellos dentro del término estipulado los perderá, sin poder hacer uso de ellos posteriormente.'
  ]);
  pBeneficio.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  
  // VII. CONDICIONES Y RESTRICCIONES
  appendRichSection(body, 'VII. Condiciones y Restricciones: ', [
    'Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:'
  ]);
  
  const restricciones = [
    'Campaña válida únicamente para órdenes realizadas a través de {{REF_TIENDA}}, al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.',
    'La presente Campaña se encuentra sujeta a los horarios de operación de los puntos de venta de {{REF_TIENDA}}, que ofrecen y exhiben sus productos al interior de la Plataforma Rappi.',
    'Los descuentos de los productos y/o servicios objeto de la Campaña no son intercambiables ni transferibles.',
    'Campaña válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.',
    'Campaña válida durante la Vigencia y/o hasta agotar existencias, lo que primero ocurra.',
    'El Beneficio obtenido en virtud de la presente Campaña no es acumulable con otras promociones exhibidas en la Plataforma Rappi.'
  ];
  restricciones.forEach(r => appendRichListItem(body, r));
  
  appendRichListItem(body, ['Máximo ', {text: '{{LIMITE_ORDENES}}', bold: true}, ' {{TEXTO_ORDENES}} por Usuario/Consumidor.']);
  
  appendRichListItem(body, ['Se aclara que el monto máximo de Créditos a recibir por el Usuario/Consumidor Participante es de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true},
    '), de acuerdo con lo indicado en la sección VI (Beneficio) de los presentes Términos y Condiciones.']);
  
  appendRichListItem(body, 'En caso de cancelación total o parcial de la orden, el Usuario/Consumidor Participante no tendrá derecho al Beneficio.');
  appendRichListItem(body, '{{TEXTO_LUGAR_REDENCION}}');
  appendRichListItem(body, 'Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero y la finalidad de su redención no es recibir una cantidad de dinero en efectivo.');
  appendRichListItem(body, 'Se aclara que los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni tampoco pueden ser utilizados para pagar el valor del costo de envío, la tarifa de servicio de un pedido realizado a través de la Plataforma Rappi y/o la propina otorgada de forma voluntaria a los repartidores independientes.');
  
  // VIII. MEDIO DE PAGO
  appendRichSection(body, 'VIII. Medio de Pago: ', ['{{TEXTO_METODO_PAGO}}']);
  
  // IX. MODIFICACIONES
  appendRichSection(body, 'IX. Modificaciones e Interpretación: ', [
    'Rappi se reserva el derecho de cancelar órdenes si detecta un comportamiento irregular por parte del Usuario/Consumidor en la Plataforma Rappi. Rappi se reserva el derecho de rechazar y cancelar cualquier orden que, por sus características, Rappi determine que no aplica para el Beneficio de la presente Campaña, sin previo aviso al Usuario/Consumidor.'
  ]);
  
  // X. DECLARACIÓN
  appendRichSection(body, 'X. Declaración: ', [
    'El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi {{DECLARACION_TIENDA}}. Rappi no comercializa productos puesto que es solo una plataforma tecnológica de contacto.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, los cuales se encuentran en la siguiente dirección electrónica: https://legal.rappi.com.co/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.'
  ]);
  
  // XI. JURISDICCIÓN
  appendRichSection(body, 'XI. Jurisdicción y Solución de conflictos: ', [
    'Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Campaña o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.'
  ]);
  
  doc.saveAndClose();
  Logger.log('📄 Template Cashback creado: ' + doc.getId());
  return doc.getId();
}

// ---- CONCURSO TEMPLATE ----
function seedConcursoTemplate() {
  const doc = DocumentApp.create('Template_CO_Concurso');
  const body = doc.getBody();
  
  // TÍTULO
  const titulo = body.appendParagraph('TÉRMINOS Y CONDICIONES – ACTIVIDAD PROMOCIONAL "{{NOMBRE_CAMPANA_UPPER}}"');
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  
  // INTRODUCCIÓN
  appendRichParagraph(body, [
    'Por medio del presente documento se dan a conocer los términos y condiciones de la actividad promocional denominada "',
    {text: '{{NOMBRE_CAMPANA}}', bold: true},
    '" (en adelante la "Actividad Promocional"). La sola participación en esta actividad implica el conocimiento y aceptación total e incondicional de estos Términos y Condiciones, no pudiendo alegar el participante el desconocimiento de estos.'
  ]);
  
  // I. ORGANIZADOR Y TERRITORIO
  appendRichSection(body, 'I. Organizador y Territorio: ', [
    'La presente Actividad Promocional es organizada y financiada exclusivamente por ',
    {text: '{{ORGANIZADOR}}', bold: true},
    ' (en adelante el "Organizador"), quien es el responsable de la entrega del beneficio. RAPPI S.A.S. actúa únicamente como plataforma de contacto y medio de difusión. La Actividad Promocional será válida únicamente para pedidos realizados a través de la Plataforma Rappi en las zonas de cobertura habilitadas en ',
    {text: '{{TEXTO_TERRITORIO}}', bold: true},
    ' (en adelante el "Territorio").'
  ]);
  
  // II. VIGENCIA
  appendRichSection(body, 'II. Vigencia: ', [
    'La Actividad Promocional estará vigente desde las ',
    {text: '{{HORA_INICIO}}', bold: true}, ' del ', {text: '{{FECHA_INICIO}}', bold: true},
    ' hasta las ', {text: '{{HORA_FIN}}', bold: true}, ' del ', {text: '{{FECHA_FIN}}', bold: true},
    ' (en adelante, la "Vigencia"). Los pedidos realizados fuera de este periodo no serán tenidos en cuenta.'
  ]);
  
  // III. PARTICIPANTES
  appendRichSection(body, 'III. Participantes: ', ['{{TEXTO_SEGMENTO}} Adicionalmente, para participar se debe:']);
  appendRichListItem(body, '(i) Ser mayor de edad y residir en el Territorio.');
  appendRichListItem(body, '(ii) Tener una cuenta activa y vigente en la Plataforma Rappi.');
  appendRichListItem(body, '(iii) Realizar compras de los Productos Participantes durante la Vigencia, cumpliendo con la mecánica descrita.');
  
  // IV. TIENDA Y PRODUCTOS
  appendRichSection(body, 'IV. Tienda y Productos Participantes: ', [
    'Participará la tienda virtual ', {text: '"{{TIENDA_BASE}}"', bold: true},
    ' (en adelante la "Tienda Participante") al interior de la Plataforma Rappi, ubicada dentro del Territorio. Participarán ',
    {text: '{{PRODUCTOS_PARTICIPANTES}}', bold: true}, ' disponibles en las secciones/verticales de ',
    {text: '{{VERTICALES}}', bold: true}, ' (en adelante, los "Productos Participantes").'
  ]);
  
  // V. MECÁNICA
  appendRichSection(body, 'V. Mecánica de la Actividad Promocional: ', [
    'La presente actividad es un concurso de destreza comercial. Serán seleccionados como ',
    {text: '{{PLURAL_GANADORES}}', bold: true}, ' los ',
    {text: '{{NUM_GANADORES_LETRAS}} ({{NUM_GANADORES}})', bold: true},
    ' Usuarios/Consumidores que, durante la Vigencia, registren el ',
    {text: '{{CRITERIO_GANADOR}}', bold: true},
    ' de Productos Participantes en la Tienda Participante.'
  ]);
  
  appendRichParagraph(body, ['Condiciones de participación:']);
  appendRichListItem(body, 'Solo se tendrán en cuenta órdenes finalizadas y entregadas. No suman órdenes canceladas, devueltas o con contracargo.');
  appendRichListItem(body, 'Compra Mínima por orden: {{MINIMO_COMPRA_TEXTO}}');
  appendRichListItem(body, 'No se tendrán en cuenta pagos realizados parcialmente con RappiCréditos, ni órdenes manipuladas fraudulentamente.');
  
  // VI. DESEMPATE
  appendRichSection(body, 'VI. Criterios de Desempate: ', [
    'En caso de presentarse un empate entre dos o más participantes, se definirá el ganador aplicando los siguientes criterios en estricto orden:'
  ]);
  appendRichListItem(body, '{{DESEMPATE_1}}');
  appendRichListItem(body, 'Si persiste el empate, ganará el Usuario/Consumidor que haya alcanzado la cifra registrada primero en el tiempo, según los registros del sistema de la Plataforma Rappi (criterio cronológico).');
  
  // VII. PREMIO
  appendRichSection(body, 'VII. Premio: ', [
    'Se entregarán ', {text: '{{NUM_GANADORES_LETRAS}} ({{NUM_GANADORES}})', bold: true}, ' premios consistentes en:'
  ]);
  appendRichListItem(body, {text: '{{LISTA_PREMIOS}}', bold: true});
  
  // Bloque créditos (se limpia si es premio físico)
  appendRichParagraph(body, [
    'Los Créditos serán cargados a la cuenta del ganador dentro de los ',
    {text: '{{DIAS_CARGA_LETRAS}} ({{DIAS_CARGA_NUM}})', bold: true},
    ' días hábiles siguientes a la confirmación del ganador. Los Créditos tendrán una vigencia de ',
    {text: '{{VIGENCIA_CREDITOS_LETRAS}} ({{VIGENCIA_CREDITOS_NUM}})', bold: true},
    ' días calendario contados a partir de su carga.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los Créditos otorgados podrán ser utilizados/redimidos ',
    {text: '{{LUGAR_REDENCION_PREMIO}}', bold: true}, ', únicamente dentro del Territorio.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero. Los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni para pagar el costo de envío, tarifa de servicio o propina.'
  ]);
  
  // VIII. CONDICIONES
  appendRichSection(body, 'VIII. Condiciones y Restricciones: ', [
    'Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:'
  ]);
  
  appendRichListItem(body, 'Actividad válida únicamente para órdenes realizadas a través de la Tienda Participante, al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.');
  appendRichListItem(body, 'La presente Actividad Promocional se encuentra sujeta a los horarios de operación de los puntos de venta de la Tienda Participante.');
  appendRichListItem(body, 'Actividad válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.');
  appendRichListItem(body, 'Actividad válida durante la Vigencia.');
  appendRichListItem(body, 'El Beneficio obtenido en virtud de la presente Actividad Promocional no es acumulable con otras promociones exhibidas en la Plataforma Rappi.');
  appendRichListItem(body, 'En caso de cancelación total o parcial de la orden, dicha orden no sumará al acumulado del Usuario/Consumidor.');
  
  // IX. MEDIO DE PAGO
  appendRichSection(body, 'IX. Medio de Pago: ', ['{{TEXTO_METODO_PAGO}}']);
  
  // X. ANUNCIO
  appendRichSection(body, 'X. Anuncio y Entrega de Premios: ', [
    'Los ganadores serán anunciados el día ', {text: '{{FECHA_ANUNCIO}}', bold: true},
    ' a través de los canales oficiales definidos por el Organizador (con el apoyo de difusión de la Plataforma Rappi, de ser requerido).'
  ]);
  appendRichListItem(body, 'La gestión, logística, costos de envío y entrega efectiva de los premios correrán por cuenta y responsabilidad exclusiva de {{RESPONSABLE_ENTREGA}}.');
  appendRichListItem(body, 'El ganador tendrá un plazo máximo de cinco (5) días hábiles para responder al contacto. Si no responde en dicho plazo, se entenderá que renuncia al premio y se procederá a contactar al siguiente participante en el ranking.');
  
  // XI. EXCLUSIONES
  appendRichSection(body, 'XI. Exclusiones y Fraude: ', [
    'El Organizador y Rappi se reservan el derecho de excluir a cualquier participante y cancelar la entrega del premio si detectan:'
  ]);
  appendRichListItem(body, '(i) Compras inusuales, autocompras, fraude, o uso de cuentas múltiples.');
  appendRichListItem(body, '(ii) Pagos disputados o contracargos.');
  appendRichListItem(body, '(iii) Violación a los Términos y Condiciones generales de la Plataforma Rappi.');
  appendRichListItem(body, '(iv) Cualquier comportamiento irregular que atente contra la naturaleza de la Actividad Promocional.');
  
  // XII. DECLARACIÓN
  appendRichSection(body, 'XII. Declaración sobre el Rol de Rappi: ', [
    'El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi es la Tienda Participante (Aliado Comercial). Rappi no comercializa productos, ya que es una plataforma de contacto. Rappi actúa únicamente como medio de difusión y comunicación de material publicitario para la realización de la Actividad Promocional.'
  ]);
  
  // XIII. LIMITACIÓN
  appendRichSection(body, 'XIII. Limitación de Responsabilidad: ', [
    'La actividad se brinda como un medio de esparcimiento y ocio para el público en general. El Organizador y/o Rappi no asumen responsabilidad por ninguna consecuencia que resulte directa o indirectamente de cualquier acción o falta de acción que el participante emprenda.'
  ]);
  
  // XIV. DATOS
  appendRichSection(body, 'XIV. Tratamiento de Datos Personales: ', [
    'El participante autoriza el tratamiento de sus datos personales para fines de la actividad. Si la entrega del premio requiere de un tercero (el Organizador), Rappi transferirá los datos de contacto necesarios bajo un acuerdo de transmisión de datos seguro, conforme a la Ley 1581 de 2012.'
  ]);
  
  // XV. CONTACTO
  appendRichSection(body, 'XV. Contacto y Atención de PQR: ', [
    'Cualquier duda o inquietud sobre los alcances e interpretación de los presentes Términos y Condiciones, podrá ser consultada a través de los siguientes canales de contacto del Organizador:'
  ]);
  appendRichListItem(body, ['Teléfono de Atención: ', {text: '{{TELEFONO_CONTACTO}}', bold: true}]);
  appendRichListItem(body, ['Correo electrónico: ', {text: '{{EMAIL_CONTACTO}}', bold: true}]);
  appendRichParagraph(body, ['Para dudas relacionadas con el funcionamiento de la Plataforma Rappi, el Usuario podrá contactar al Centro de Ayuda de Rappi a través de la aplicación.']);
  
  // XVI. MODIFICACIONES
  appendRichSection(body, 'XVI. Modificaciones: ', [
    'El Organizador se reserva el derecho de modificar los presentes Términos y Condiciones, así como suspender o cancelar la Actividad Promocional por motivos de fuerza mayor o caso fortuito, informando previamente a los usuarios/consumidores y sin perjuicio de los derechos adquiridos de los participantes.'
  ]);
  
  // XVII. JURISDICCIÓN
  appendRichSection(body, 'XVII. Jurisdicción y Solución de Conflictos: ', [
    'Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Actividad Promocional o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, disponibles en: https://legal.rappi.com/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.'
  ]);
  
  doc.saveAndClose();
  Logger.log('📄 Template Concurso creado: ' + doc.getId());
  return doc.getId();
}

// ---- SEED REGISTRY ----
function _seedRegistry(cashbackDocId, concursoDocId) {
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET_NAME);
    const headers = ['country_code', 'country_name', 'campaign_type', 'template_doc_id', 'version', 'status', 'currency_code', 'currency_symbol', 'legal_owner', 'last_updated', 'notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  const today = new Date().toISOString().split('T')[0];
  
  sheet.appendRow(['CO', 'Colombia', 'Cashback', cashbackDocId, '1.0', 'active', 'COP', '$', 'juan.gallego@rappi.com', today, 'Template base generado automáticamente']);
  sheet.appendRow(['CO', 'Colombia', 'Concurso Mayor Comprador', concursoDocId, '1.0', 'active', 'COP', '$', 'juan.gallego@rappi.com', today, 'Template base generado automáticamente']);
}

// ---- SEED FIELDS ----
function seedColombiaFields() {
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(FIELDS_SHEET_NAME);
    const headers = ['field_id', 'country_code', 'campaign_type', 'placeholder', 'label_es', 'field_type', 'icon', 'required', 'section', 'validation_rule', 'options', 'default_value', 'tooltip', 'depends_on', 'order', 'group'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  if (sheet.getLastRow() > 1) {
    Logger.log('⚠️ Template_Fields ya tiene datos. Saltando seed.');
    return;
  }
  
  const fields = [
    ['userEmail', 'ALL', 'ALL', '', 'Email corporativo', 'email', 'fa-envelope', 'TRUE', '1', '', '', '', '', '', '1', 'Info Básica'],
    ['dynamicType', 'ALL', 'ALL', '', 'Tipo de campaña', 'select', 'fa-layer-group', 'TRUE', '1', '', 'Cashback|Concurso Mayor Comprador', '', '', '', '2', 'Info Básica'],
    ['campaignName', 'ALL', 'ALL', '{{NOMBRE_CAMPANA}}', 'Nombre campaña', 'text', 'fa-tag', 'FALSE', '1', '', '', '', 'Se genera automático si vacío', '', '3', 'Info Básica'],
    ['shopName', 'ALL', 'ALL', '{{TIENDA_BASE}}', 'Nombre tienda', 'text', 'fa-shop', 'TRUE', '2', '', '', '', '', '', '1', 'Aliado'],
    ['territory', 'ALL', 'ALL', '{{TEXTO_TERRITORIO}}', 'Territorio', 'text', 'fa-map-location-dot', 'TRUE', '2', '', '', '', '', '', '2', 'Aliado'],
    ['startDate', 'ALL', 'ALL', '{{FECHA_INICIO}}', 'Fecha inicio', 'date', 'fa-calendar', 'TRUE', '2', '', '', '', '', '', '3', 'Vigencia'],
    ['endDate', 'ALL', 'ALL', '{{FECHA_FIN}}', 'Fecha fin', 'date', 'fa-calendar-check', 'TRUE', '2', '', '', '', '', '', '4', 'Vigencia'],
    ['cashbackPct', 'ALL', 'Cashback', '{{TEXTO_PORCENTAJE}}', '% Cashback', 'number', 'fa-percent', 'TRUE', '3', 'min:1,max:100', '', '', '', '', '1', 'Cashback'],
    ['cap', 'ALL', 'Cashback', '{{TOPE_NUM}}', 'Tope máximo', 'number', 'fa-ban', 'TRUE', '3', '', '', '', '', '', '2', 'Cashback'],
    ['budget', 'ALL', 'Cashback', '{{PRESUPUESTO_NUM}}', 'Presupuesto', 'number', 'fa-sack-dollar', 'TRUE', '3', '', '', '', '', '', '3', 'Cashback'],
    ['redemptionPlace', 'ALL', 'Cashback', '{{TEXTO_LUGAR_REDENCION}}', 'Lugar redención', 'select', 'fa-location-crosshairs', 'FALSE', '3', '', 'Brand Credits|Restaurantes|Créditos Generales', '', '', '', '4', 'Cashback'],
    ['organizerLegalName', 'ALL', 'Concurso Mayor Comprador', '{{ORGANIZADOR}}', 'Razón social', 'text', 'fa-briefcase', 'TRUE', '3', '', '', '', '', '', '1', 'Organizador'],
    ['organizerPhone', 'ALL', 'Concurso Mayor Comprador', '{{TELEFONO_CONTACTO}}', 'Teléfono', 'text', 'fa-phone', 'TRUE', '3', '', '', '', '', '', '2', 'Organizador'],
    ['organizerEmail', 'ALL', 'Concurso Mayor Comprador', '{{EMAIL_CONTACTO}}', 'Email PQR', 'email', 'fa-envelope', 'TRUE', '3', '', '', '', '', '', '3', 'Organizador'],
    ['numberOfWinners', 'ALL', 'Concurso Mayor Comprador', '{{NUM_GANADORES}}', '# Ganadores', 'number', 'fa-users', 'TRUE', '3', '', '', '1', '', '', '4', 'Mecánica'],
    ['winnerCriteria', 'ALL', 'Concurso Mayor Comprador', '{{CRITERIO_GANADOR}}', 'Criterio ganador', 'select', 'fa-filter', 'TRUE', '3', '', 'Mayor Venta ($)|Más Órdenes (#)', '', '', '', '5', 'Mecánica'],
    ['announcementDate', 'ALL', 'Concurso Mayor Comprador', '{{FECHA_ANUNCIO}}', 'Fecha anuncio', 'date', 'fa-bullhorn', 'TRUE', '3', '', '', '', '', '', '6', 'Mecánica'],
    ['paymentMethods', 'ALL', 'ALL', '{{TEXTO_METODO_PAGO}}', 'Métodos de pago', 'select', 'fa-credit-card', 'FALSE', '4', '', 'Todos excepto Efectivo|Todos|Únicamente RappiCard', 'Todos excepto Efectivo', '', '', '1', 'Restricciones'],
    ['userSegment', 'ALL', 'ALL', '{{TEXTO_SEGMENTO}}', 'Segmento usuarios', 'select', 'fa-users', 'FALSE', '4', '', 'Todos los usuarios|Pro y Pro Black|Nuevos Usuarios|Reactivos', 'Todos los usuarios', '', '', '2', 'Restricciones'],
    ['maxOrders', 'ALL', 'ALL', '{{LIMITE_ORDENES}}', 'Máx órdenes', 'text', 'fa-list-ol', 'FALSE', '4', '', '', '1', '', '', '3', 'Restricciones'],
    ['specialConditions', 'ALL', 'ALL', '{{CONDICIONES_ESPECIALES}}', 'Condiciones especiales', 'textarea', 'fa-exclamation-circle', 'FALSE', '4', '', '', '', '', '', '5', 'Restricciones']
  ];
  
  fields.forEach(f => sheet.appendRow(f));
  Logger.log('✅ ' + fields.length + ' campos creados en Template_Fields');
}

// =================================================================
// 3. CARPETAS ESTRUCTURADAS EN DRIVE
// =================================================================
function setupTemplateFolders() {
  Logger.log('📁 Creando estructura de carpetas...');

  // Crear carpeta raíz
  const root = _getOrCreateDriveFolder(null, TEMPLATES_ROOT_NAME);
  Logger.log('📁 Raíz: ' + root.getName() + ' (' + root.getId() + ')');

  const countries = [
    { code: 'CO', name: 'Colombia' },
    { code: 'MX', name: 'México' },
    { code: 'CR', name: 'Costa Rica' },
    { code: 'BR', name: 'Brasil' },
    { code: 'AR', name: 'Argentina' },
    { code: 'UY', name: 'Uruguay' },
    { code: 'CL', name: 'Chile' },
    { code: 'EC', name: 'Ecuador' }
  ];

  const types = ['Cashback', 'Concurso'];

  countries.forEach(country => {
    const countryFolder = _getOrCreateDriveFolder(root, `${country.code}_${country.name}`);
    types.forEach(type => {
      _getOrCreateDriveFolder(countryFolder, type);
    });
  });

  // Carpetas especiales
  _getOrCreateDriveFolder(root, '_Borradores');
  _getOrCreateDriveFolder(root, '_Archivo');

  Logger.log('✅ Estructura de carpetas creada');
  Logger.log('📋 Folder ID raíz: ' + root.getId());

  // Guardar ID de carpeta raíz en propiedades del script
  PropertiesService.getScriptProperties().setProperty('TEMPLATES_FOLDER_ID', root.getId());

  return root.getId();
}

// =================================================================
// 6. SETUP ADMIN SYSTEM — EJECUTAR UNA VEZ
// =================================================================
function setupAdminSystem() {
  Logger.log('🔐 ═══════════════════════════════════════');
  Logger.log('🔐 RAPPIMIND ADMIN SYSTEM — SETUP');
  Logger.log('🔐 ═══════════════════════════════════════');

  // ──── PASO 1: Crear sheet de equipo ────
  Logger.log('👥 Creando sheet de equipo...');
  const teamSheet = _getOrCreateSheet(TEAM_SHEET_NAME,
    ['email', 'name', 'role', 'added_by', 'added_date', 'status', 'notes']);

  // Seed: Juan como owner
  const callerEmail = Session.getActiveUser().getEmail();
  const existingTeam = _sheetToObjects(teamSheet);
  if (existingTeam.length === 0) {
    teamSheet.appendRow([callerEmail, 'Juan (Owner)', 'owner', 'system', new Date().toISOString().split('T')[0], 'active', 'Setup inicial']);
    Logger.log('✅ ' + callerEmail + ' agregado como owner');
  }

  // ──── PASO 2: Agregar columnas de aprobación al Registry ────
  Logger.log('📋 Actualizando Registry con columnas de aprobación...');
  _upgradeRegistryColumns();

  // ──── PASO 3: Crear approval log ────
  Logger.log('📝 Creando log de aprobaciones...');
  _getOrCreateSheet(APPROVAL_LOG_SHEET,
    ['timestamp', 'actor', 'action', 'details']);

  // ──── PASO 4: Crear carpetas ────
  Logger.log('📁 Creando estructura de carpetas...');
  const rootId = setupTemplateFolders();

  // ──── PASO 5: Mover templates existentes a carpetas ────
  Logger.log('📦 Moviendo templates existentes...');
  _moveExistingTemplatesToFolders();

  Logger.log('');
  Logger.log('🎉 ═══════════════════════════════════════');
  Logger.log('🎉 ADMIN SYSTEM CONFIGURADO');
  Logger.log('🎉 ═══════════════════════════════════════');
  Logger.log('');
  Logger.log('👤 Owner: ' + callerEmail);
  Logger.log('📁 Carpeta raíz: ' + rootId);
  Logger.log('');
  Logger.log('👉 SIGUIENTE: Agrega miembros al equipo desde el Admin Panel');
  Logger.log('   o manualmente en la hoja Admin_Team');
}

function _upgradeRegistryColumns() {
  const sheet = _getSheet(REGISTRY_SHEET_NAME);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newCols = ['submitted_by', 'submitted_date', 'approved_by', 'approved_date', 'rejected_by'];

  newCols.forEach(col => {
    if (headers.indexOf(col) === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col)
        .setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
      Logger.log('  + Columna: ' + col);
    }
  });
}

function _moveExistingTemplatesToFolders() {
  try {
    const registry = _getTemplateRegistry();
    registry.forEach(tpl => {
      if (tpl.template_doc_id) {
        _moveTemplateToFolder(
          tpl.template_doc_id,
          tpl.country_code,
          tpl.country_name || tpl.country_code,
          tpl.campaign_type
        );
      }
    });
  } catch (e) {
    Logger.log('⚠️ Error moviendo templates: ' + e.message);
  }
}

function setupCampaignTypes() {
  Logger.log('🎯 Creando catálogo de dinámicas...');
  
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(CAMPAIGN_TYPES_SHEET);
  
  if (sheet) {
    Logger.log('⚠️ Campaign_Types ya existe. Saltando.');
    return;
  }
  
  sheet = ss.insertSheet(CAMPAIGN_TYPES_SHEET);
  const headers = ['type_id', 'type_name', 'description', 'parent_type', 'processing_mode', 'icon', 'color', 'status', 'countries', 'created_by', 'created_date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  const callerEmail = Session.getActiveUser().getEmail();
  const today = new Date().toISOString().split('T')[0];
  
  // Sembrar tipos iniciales (legacy = usan el motor hardcodeado actual)
  const types = [
    ['cashback', 'Cashback', 'Créditos devueltos al usuario por compra en tienda participante', '', 'legacy', 'fa-coins', '#00D68F', 'active', 'ALL', callerEmail, today],
    ['concurso', 'Concurso Mayor Comprador', 'Concurso donde ganan los usuarios que más compren (valor o cantidad)', '', 'legacy', 'fa-trophy', '#8B5CF6', 'active', 'ALL', callerEmail, today]
  ];
  
  types.forEach(t => sheet.appendRow(t));
  
  Logger.log('✅ Campaign_Types creada con ' + types.length + ' tipos iniciales');
  Logger.log('');
  Logger.log('👉 Para agregar nuevos tipos, usa el Admin Panel > pestaña Dinámicas');
  Logger.log('   o agrega filas directamente a la sheet Campaign_Types');
}

function setupCountrySettings() {
  Logger.log('🌎 Creando configuración de países...');
  
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(COUNTRY_SETTINGS_SHEET);
  
  if (sheet) {
    Logger.log('⚠️ Country_Settings ya existe. Saltando.');
    return;
  }
  
  sheet = ss.insertSheet(COUNTRY_SETTINGS_SHEET);
  const headers = ['country_code', 'country_name', 'legal_country', 'currency_name', 'currency_code', 'currency_symbol', 'legal_entity', 'timezone'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  const countries = [
    ['CO', 'Colombia', 'la República de Colombia', 'pesos M/Cte.', 'COP', '$', 'la Superintendencia de Industria y Comercio (SIC)', 'America/Bogota'],
    ['MX', 'México', 'los Estados Unidos Mexicanos', 'pesos MXN', 'MXN', '$', 'la Procuraduría Federal del Consumidor (PROFECO)', 'America/Mexico_City'],
    ['PE', 'Perú', 'la República del Perú', 'soles', 'PEN', 'S/', 'el Instituto Nacional de Defensa de la Competencia y de la Protección de la Propiedad Intelectual (INDECOPI)', 'America/Lima'],
    ['CL', 'Chile', 'la República de Chile', 'pesos CLP', 'CLP', '$', 'el Servicio Nacional del Consumidor (SERNAC)', 'America/Santiago'],
    ['AR', 'Argentina', 'la República Argentina', 'pesos ARS', 'ARS', '$', 'la Dirección Nacional de Defensa del Consumidor', 'America/Buenos_Aires'],
    ['EC', 'Ecuador', 'la República del Ecuador', 'dólares USD', 'USD', '$', 'la Defensoría del Pueblo', 'America/Guayaquil'],
    ['UY', 'Uruguay', 'la República Oriental del Uruguay', 'pesos UYU', 'UYU', '$', 'el Área de Defensa del Consumidor', 'America/Montevideo'],
    ['CR', 'Costa Rica', 'la República de Costa Rica', 'colones CRC', 'CRC', '₡', 'el Ministerio de Economía, Industria y Comercio (MEIC)', 'America/Costa_Rica'],
    ['BR', 'Brasil', 'la República Federativa del Brasil', 'reais BRL', 'BRL', 'R$', 'el PROCON y la Secretaría Nacional del Consumidor (SENACON)', 'America/Sao_Paulo']
  ];
  
  countries.forEach(c => sheet.appendRow(c));
  
  Logger.log('✅ Country_Settings creada con ' + countries.length + ' países');
}

function upgradeTemplateFieldsFormatting() {
  Logger.log('📋 Actualizando Template_Fields con columna format_as...');
  
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  const sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
  if (!sheet) {
    Logger.log('⚠️ Template_Fields no existe aún. Se creará cuando se necesite.');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (headers.indexOf('format_as') === -1) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol).setValue('format_as');
    sheet.getRange(1, nextCol).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    Logger.log('✅ Columna format_as agregada');
    
    // Actualizar campos existentes con format_as correcto
    const data = sheet.getDataRange().getValues();
    const fieldIdCol = headers.indexOf('field_id');
    
    // Mapeo de campos conocidos a su formato
    const formatMap = {
      'cashbackPct': 'percentage',
      'cap': 'money',
      'budget': 'money',
      'startDate': 'date_legal',
      'endDate': 'date_legal',
      'announcementDate': 'date_legal',
      'numberOfWinners': 'number_words',
      'maxOrders': 'number_words'
    };
    
    for (let i = 1; i < data.length; i++) {
      const fieldId = data[i][fieldIdCol];
      if (formatMap[fieldId]) {
        sheet.getRange(i + 1, nextCol).setValue(formatMap[fieldId]);
      }
    }
    Logger.log('✅ Formatos asignados a campos conocidos');
  } else {
    Logger.log('ℹ️ Columna format_as ya existe');
  }
}

function upgradeRegistryVertical() {
  Logger.log('📋 Actualizando Template_Registry con columna vertical...');
  
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  const sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  if (!sheet) {
    Logger.log('⚠️ Template_Registry no existe.');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (headers.indexOf('vertical') === -1) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol).setValue('vertical');
    sheet.getRange(1, nextCol).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    Logger.log('✅ Columna vertical agregada al Registry');
    
    // Setear 'ALL' en todos los templates existentes
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      for (let i = 2; i <= lastRow; i++) {
        sheet.getRange(i, nextCol).setValue('ALL');
      }
      Logger.log('✅ Templates existentes marcados como vertical: ALL');
    }
  } else {
    Logger.log('ℹ️ Columna vertical ya existe');
  }
}
function seedMissingFields() {
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  const sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
  if (!sheet) { Logger.log('❌ Template_Fields no existe. Ejecuta setupCompleto primero.'); return; }

  // Leer field_ids existentes para no duplicar
  const existing = sheet.getDataRange().getValues();
  const existingIds = existing.slice(1).map(r => r[0]);

  const newFields = [
    // Sección 2 — Campos faltantes
    ['startTime', 'ALL', 'ALL', '', 'Hora inicio', 'time', 'fa-clock', 'FALSE', '2', '', '', '12:00', '', '', '5', 'Vigencia'],
    ['endTime', 'ALL', 'ALL', '', 'Hora fin', 'time', 'fa-clock', 'FALSE', '2', '', '', '23:59', '', '', '6', 'Vigencia'],

    // Sección 3 — Cashback avanzado
    ['loadType', 'ALL', 'Cashback', '', 'Momento de carga', 'select', 'fa-clock', 'FALSE', '3', '', 'Inmediatamente (Al finalizar la orden)|Al día siguiente|Fecha específica', 'Inmediatamente (Al finalizar la orden)', '', '', '5', 'Cashback Avanzado'],
    ['loadDate', 'ALL', 'Cashback', '', 'Fecha de carga', 'date', 'fa-calendar-day', 'FALSE', '3', '', '', '', '', 'loadType:Fecha específica', '6', 'Cashback Avanzado'],
    ['validityType', 'ALL', 'Cashback', '', 'Vigencia de créditos', 'select', 'fa-hourglass', 'FALSE', '3', '', 'Por días calendario (Duración)|Por fechas específicas (Rango)', 'Por días calendario (Duración)', '', '', '7', 'Cashback Avanzado'],
    ['validityDays', 'ALL', 'Cashback', '', 'Días de vigencia', 'number', 'fa-hashtag', 'FALSE', '3', '', '', '30', '', 'validityType:Por días calendario (Duración)', '8', 'Cashback Avanzado'],
    ['redemptionStart', 'ALL', 'Cashback', '', 'Inicio redención', 'date', 'fa-calendar', 'FALSE', '3', '', '', '', '', 'validityType:Por fechas específicas (Rango)', '9', 'Cashback Avanzado'],
    ['redemptionEnd', 'ALL', 'Cashback', '', 'Fin redención', 'date', 'fa-calendar-check', 'FALSE', '3', '', '', '', '', 'validityType:Por fechas específicas (Rango)', '10', 'Cashback Avanzado'],

    // Sección 3 — Concurso (faltantes)
    ['verticals', 'ALL', 'Concurso Mayor Comprador', '{{VERTICALES}}', 'Secciones participantes', 'text', 'fa-layer-group', 'FALSE', '3', '', '', 'Restaurantes', '', '', '7', 'Mecánica'],
    ['participatingProducts', 'ALL', 'Concurso Mayor Comprador', '{{PRODUCTOS_PARTICIPANTES}}', 'Productos participantes', 'text', 'fa-box-open', 'FALSE', '3', '', '', '', '', '', '8', 'Mecánica'],
    ['prizeType', 'ALL', 'Concurso Mayor Comprador', '', 'Tipo de premio', 'select', 'fa-gift', 'TRUE', '3', '', 'credits|physical', 'credits', '', '', '9', 'Premio'],
    ['creditsAmount', 'ALL', 'Concurso Mayor Comprador', '', 'Monto créditos premio', 'number', 'fa-coins', 'FALSE', '3', '', '', '', '', 'prizeType:credits', '10', 'Premio'],
    ['creditLoadDays', 'ALL', 'Concurso Mayor Comprador', '', 'Días para carga premio', 'number', 'fa-clock', 'FALSE', '3', '', '', '5', '', 'prizeType:credits', '11', 'Premio'],
    ['creditsValidityDays', 'ALL', 'Concurso Mayor Comprador', '', 'Vigencia créditos premio', 'number', 'fa-hourglass-half', 'FALSE', '3', '', '', '30', '', 'prizeType:credits', '12', 'Premio'],
    ['creditsRedemptionPlace', 'ALL', 'Concurso Mayor Comprador', '', 'Lugar redención premio', 'select', 'fa-location-dot', 'FALSE', '3', '', 'Únicamente en la Tienda Participante|En cualquier tienda de Restaurantes|En cualquier sección de la Plataforma', 'Únicamente en la Tienda Participante', '', 'prizeType:credits', '13', 'Premio'],
    ['physicalPrizeDescription', 'ALL', 'Concurso Mayor Comprador', '', 'Descripción premio físico', 'textarea', 'fa-gift', 'FALSE', '3', '', '', '', '', 'prizeType:physical', '14', 'Premio'],
    ['prizeDeliveryBy', 'ALL', 'Concurso Mayor Comprador', '', 'Quién entrega premio', 'select', 'fa-truck', 'FALSE', '3', '', 'rappi|organizer', 'rappi', '', 'prizeType:physical', '15', 'Premio'],
    ['minParticipation', 'ALL', 'Concurso Mayor Comprador', '', 'Mínimo compra participar', 'text', 'fa-dollar-sign', 'FALSE', '3', '', '', '', '', '', '16', 'Mecánica'],

    // Sección 4 — Faltantes
    ['minPurchase', 'ALL', 'Cashback', '', 'Mínimo de compra', 'number', 'fa-basket-shopping', 'FALSE', '4', '', '', '', '', '', '0', 'Restricciones'],
    ['extraEmails', 'ALL', 'ALL', '', 'Emails CC', 'text', 'fa-paper-plane', 'FALSE', '4', '', '', '', 'Separados por coma', '', '4', 'Restricciones'],
    ['specialConditions', 'ALL', 'ALL', '{{CONDICIONES_ESPECIALES}}', 'Condiciones Especiales', 'select', 'fa-exclamation-circle', 'FALSE', '4', '', 'Sin condiciones especiales|Aplica únicamente para consumo en el local (Dine-in)|Aplica únicamente para órdenes con envío a domicilio|Campaña válida únicamente para las primeras 100 órdenes del día|No aplica para combos ni promociones, únicamente productos a precio regular', '', 'Texto legal extra', '', '1', 'Restricciones'],
  ];

  let added = 0;
  newFields.forEach(f => {
    if (!existingIds.includes(f[0])) {
      sheet.appendRow(f);
      added++;
    }
  });
  Logger.log('✅ Se agregaron ' + added + ' campos nuevos a Template_Fields');
}
