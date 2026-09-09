(() => {
  'use strict';

  const button = document.getElementById('exportEmployer');
  const yearFilter = document.getElementById('yearFilter');
  const vehicleSelect = document.getElementById('vehicleSelect');
  if (!button || !yearFilter || !vehicleSelect) return;

  const API_BASE = './api';
  const FULL_REGISTRATION_FROM = '2027-01-01';
  const MONTHS = [
    ['01', 'Januari'], ['02', 'Februari'], ['03', 'Maart'], ['04', 'April'],
    ['05', 'Mei'], ['06', 'Juni'], ['07', 'Juli'], ['08', 'Augustus'],
    ['09', 'September'], ['10', 'Oktober'], ['11', 'November'], ['12', 'December']
  ];

  const encoder = new TextEncoder();

  function xmlEscape(value) {
    return String(value ?? '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;');
  }

  function colName(index) {
    let value = index + 1;
    let result = '';
    while (value > 0) {
      value -= 1;
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  function textCell(row, col, value, style = 0) {
    if (value === '' || value === null || value === undefined) return '';
    const ref = `${colName(col)}${row}`;
    const preserve = /^\s|\s$/.test(String(value)) ? ' xml:space="preserve"' : '';
    return `<c r="${ref}" t="inlineStr" s="${style}"><is><t${preserve}>${xmlEscape(value)}</t></is></c>`;
  }

  function numberCell(row, col, value, style = 0) {
    if (value === '' || value === null || value === undefined || !Number.isFinite(Number(value))) return '';
    return `<c r="${colName(col)}${row}" s="${style}"><v>${Number(value)}</v></c>`;
  }

  function rowXml(row, cells, height = null) {
    const heightAttrs = height ? ` ht="${height}" customHeight="1"` : '';
    return `<row r="${row}"${heightAttrs}>${cells.join('')}</row>`;
  }

  function formatDate(value) {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
    return match ? `${match[3]}-${match[2]}-${match[1]}` : String(value || '');
  }

  function normalizeRides(payload) {
    return Array.isArray(payload?.rides) ? payload.rides : [];
  }

  function normalizeVehicles(payload) {
    return Array.isArray(payload?.vehicles) ? payload.vehicles : [];
  }

  async function api(path) {
    const response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      credentials: 'same-origin',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) throw new Error(`API-fout HTTP ${response.status}`);
    return response.json();
  }

  function odometers(rides) {
    return rides.flatMap((ride) => [Number(ride.startOdometer), Number(ride.endOdometer)]).filter(Number.isFinite);
  }

  function monthStats(allVehicleRides, monthRides, vehicle, year, month) {
    const businessRides = monthRides.filter((ride) => ride.type === 'business');
    const business = businessRides.reduce((sum, ride) => sum + Number(ride.distance || 0), 0);
    const monthOdometers = odometers(monthRides);
    const endOdometer = monthOdometers.length ? Math.max(...monthOdometers) : null;

    if (`${year}-${month}-01` >= FULL_REGISTRATION_FROM) {
      const privateKm = monthRides
        .filter((ride) => ride.type === 'private')
        .reduce((sum, ride) => sum + Number(ride.distance || 0), 0);
      return { business, privateKm, endOdometer, businessRides };
    }

    if (endOdometer === null) return { business, privateKm: 0, endOdometer, businessRides };

    const monthStart = `${year}-${month}-01`;
    const previousRides = allVehicleRides.filter((ride) => String(ride.date) < monthStart);
    const previousOdometers = odometers(previousRides);
    const baseline = previousOdometers.length
      ? Math.max(...previousOdometers)
      : Number(vehicle.initialOdometer);
    const privateKm = Number.isFinite(baseline) ? Math.max(0, endOdometer - baseline - business) : 0;
    return { business, privateKm, endOdometer, businessRides };
  }

  function sheetXml(vehicle, allVehicleRides, year, month) {
    const monthRides = allVehicleRides
      .filter((ride) => String(ride.date).startsWith(`${year}-${month}-`))
      .sort((a, b) => String(a.date).localeCompare(String(b.date)) || Number(a.id || 0) - Number(b.id || 0));
    const stats = monthStats(allVehicleRides, monthRides, vehicle, year, month);
    const rows = [];

    rows.push(rowXml(1, [textCell(1, 0, 'Rittenregistratie', 1), textCell(1, 9, 'Eind kilometerstand', 3)], 22));
    rows.push(rowXml(2, [textCell(2, 0, 'Jaartal', 2), textCell(2, 2, year, 5), numberCell(2, 9, stats.endOdometer, 9)]));
    rows.push(rowXml(4, [
      textCell(4, 0, 'Merk', 2), textCell(4, 2, vehicle.make || '', 5),
      textCell(4, 4, 'Type', 2), textCell(4, 6, vehicle.model || '', 5),
      textCell(4, 9, 'Totaal zakelijk gereden', 3)
    ]));
    rows.push(rowXml(5, [textCell(5, 0, 'Kenteken', 2), textCell(5, 2, vehicle.plate || '', 5), numberCell(5, 9, stats.business, 9)]));
    rows.push(rowXml(7, [
      textCell(7, 0, 'Personeelsnummer', 2), textCell(7, 2, ' ', 5),
      textCell(7, 4, 'Naam', 2), textCell(7, 6, ' ', 5), textCell(7, 9, 'Prive', 3)
    ]));
    rows.push(rowXml(8, [textCell(8, 0, 'Team manager', 2), textCell(8, 2, ' ', 5), numberCell(8, 9, stats.privateKm, 9)]));

    rows.push(rowXml(10, [
      textCell(10, 0, 'Datum', 4),
      textCell(10, 1, 'Tijd dienstrit', 4),
      textCell(10, 2, 'Begin kilometerstand', 4),
      textCell(10, 3, 'Eind kilometerstand', 4),
      textCell(10, 4, 'Gereden zakelijke kilometers', 4),
      textCell(10, 5, 'Adres van vertrek', 4),
      textCell(10, 6, 'Adres van aankomst', 4),
      textCell(10, 7, 'Opmerking', 4)
    ], 52));

    let row = 11;
    for (const ride of stats.businessRides) {
      rows.push(rowXml(row, [
        textCell(row, 0, formatDate(ride.date), 5),
        textCell(row, 1, ride.departureTime || '', 8),
        numberCell(row, 2, ride.startOdometer, 6),
        numberCell(row, 3, ride.endOdometer, 6),
        numberCell(row, 4, ride.distance, 6),
        textCell(row, 5, ride.departureAddress || '', 5),
        textCell(row, 6, ride.arrivalAddress || '', 5),
        textCell(row, 7, ride.notes || '', 5)
      ], 24));
      row += 1;
    }

    while (row <= 35) {
      rows.push(rowXml(row, [
        textCell(row, 0, ' ', 5), textCell(row, 1, ' ', 5), textCell(row, 2, ' ', 5),
        textCell(row, 3, ' ', 5), textCell(row, 4, ' ', 5), textCell(row, 5, ' ', 5),
        textCell(row, 6, ' ', 5), textCell(row, 7, ' ', 5)
      ], 22));
      row += 1;
    }

    const dimensionEnd = `K${Math.max(35, row - 1)}`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<dimension ref="A1:${dimensionEnd}"/>` +
      `<sheetViews><sheetView workbookViewId="0" showGridLines="1"><pane ySplit="10" topLeftCell="A11" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>` +
      `<col min="1" max="1" width="14" customWidth="1"/>` +
      `<col min="2" max="2" width="14" customWidth="1"/>` +
      `<col min="3" max="5" width="18" customWidth="1"/>` +
      `<col min="6" max="7" width="34" customWidth="1"/>` +
      `<col min="8" max="8" width="24" customWidth="1"/>` +
      `<col min="9" max="9" width="4" customWidth="1"/>` +
      `<col min="10" max="10" width="24" customWidth="1"/>` +
      `<col min="11" max="11" width="4" customWidth="1"/>` +
      `</cols>` +
      `<sheetData>${rows.join('')}</sheetData>` +
      `<mergeCells count="21">` +
      `<mergeCell ref="A1:H1"/><mergeCell ref="A2:B2"/><mergeCell ref="C2:D2"/>` +
      `<mergeCell ref="A4:B4"/><mergeCell ref="C4:D4"/><mergeCell ref="E4:F4"/><mergeCell ref="G4:H4"/>` +
      `<mergeCell ref="A5:B5"/><mergeCell ref="C5:D5"/>` +
      `<mergeCell ref="A7:B7"/><mergeCell ref="C7:D7"/><mergeCell ref="E7:F7"/><mergeCell ref="G7:H7"/>` +
      `<mergeCell ref="A8:B8"/><mergeCell ref="C8:H8"/>` +
      `<mergeCell ref="J1:K1"/><mergeCell ref="J2:K2"/><mergeCell ref="J4:K4"/><mergeCell ref="J5:K5"/><mergeCell ref="J7:K7"/><mergeCell ref="J8:K8"/>` +
      `</mergeCells>` +
      `<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>` +
      `<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>` +
      `</worksheet>`;
  }

  function stylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="12"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF000000"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="bottom"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function workbookXml() {
    const sheets = MONTHS.map(([, name], index) =>
      `<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>` +
      `<sheets>${sheets}</sheets><calcPr calcId="191029"/></workbook>`;
  }

  function workbookRelsXml() {
    const sheets = MONTHS.map(([,], index) =>
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets}` +
      `<Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `</Relationships>`;
  }

  function contentTypesXml() {
    const sheets = MONTHS.map(([,], index) =>
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      sheets + `</Types>`;
  }

  function rootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`;
  }

  function crc32(bytes) {
    let crc = 0xFFFFFFFF;
    for (const byte of bytes) {
      crc ^= byte;
      for (let i = 0; i < 8; i += 1) crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function u16(value) {
    return Uint8Array.of(value & 255, (value >>> 8) & 255);
  }

  function u32(value) {
    return Uint8Array.of(value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255);
  }

  function concat(parts) {
    const size = parts.reduce((sum, part) => sum + part.length, 0);
    const result = new Uint8Array(size);
    let offset = 0;
    for (const part of parts) {
      result.set(part, offset);
      offset += part.length;
    }
    return result;
  }

  function createZip(files) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const file of files) {
      const name = encoder.encode(file.name);
      const data = typeof file.data === 'string' ? encoder.encode(file.data) : file.data;
      const crc = crc32(data);
      const local = concat([
        u32(0x04034B50), u16(20), u16(0), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data
      ]);
      localParts.push(local);

      const central = concat([
        u32(0x02014B50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0),
        u16(0), u16(0), u32(0), u32(offset), name
      ]);
      centralParts.push(central);
      offset += local.length;
    }

    const central = concat(centralParts);
    const end = concat([
      u32(0x06054B50), u16(0), u16(0), u16(files.length), u16(files.length),
      u32(central.length), u32(offset), u16(0)
    ]);
    return concat([...localParts, central, end]);
  }

  function buildWorkbook(vehicle, rides, year) {
    const files = [
      { name: '[Content_Types].xml', data: contentTypesXml() },
      { name: '_rels/.rels', data: rootRelsXml() },
      { name: 'xl/workbook.xml', data: workbookXml() },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml() },
      { name: 'xl/styles.xml', data: stylesXml() }
    ];
    MONTHS.forEach(([month], index) => {
      files.push({
        name: `xl/worksheets/sheet${index + 1}.xml`,
        data: sheetXml(vehicle, rides, year, month)
      });
    });
    return createZip(files);
  }

  async function exportEmployerWorkbook() {
    const selectedVehicleId = Number(vehicleSelect.value);
    const year = String(yearFilter.value || '').trim();
    if (!Number.isInteger(selectedVehicleId) || selectedVehicleId <= 0) {
      window.alert('Kies eerst een voertuig. De werkgeversregistratie wordt per voertuig gemaakt.');
      return;
    }
    if (!/^\d{4}$/.test(year)) {
      window.alert('Kies eerst een geldig kalenderjaar.');
      return;
    }

    button.disabled = true;
    const oldText = button.textContent;
    button.textContent = 'Rittenregistratie maken…';
    try {
      const [ridesPayload, vehiclesPayload] = await Promise.all([api('/rides'), api('/vehicles')]);
      const vehicles = normalizeVehicles(vehiclesPayload);
      const vehicle = vehicles.find((item) => Number(item.id) === selectedVehicleId);
      if (!vehicle) throw new Error('Het geselecteerde voertuig bestaat niet meer.');
      const vehicleRides = normalizeRides(ridesPayload).filter((ride) => Number(ride.vehicleId) === selectedVehicleId);
      const workbook = buildWorkbook(vehicle, vehicleRides, year);
      const blob = new Blob([workbook], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `Rittenregistratie-${year}-${String(vehicle.plate || 'voertuig').replace(/[^A-Za-z0-9-]/g, '')}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Werkgeversregistratie kon niet worden gemaakt.', error);
      window.alert(`Rittenregistratie kon niet worden gemaakt: ${error.message}`);
    } finally {
      button.disabled = false;
      button.textContent = oldText;
    }
  }

  button.addEventListener('click', exportEmployerWorkbook);
})();
