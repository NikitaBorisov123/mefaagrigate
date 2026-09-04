'use strict';

importScripts('./xlsx.full.min.js');

const FORECAST_ACCURACY_THRESHOLD = 0.27;

function normalize(value) {
  return String(value == null ? '' : value).trim().toLocaleLowerCase('ru-RU').replace(/ё/g, 'е').replace(/\s+/g, ' ');
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return null;
  const parsed = Number(value.replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(parsed) ? parsed : null;
}

function toDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(Math.round((value - 25569) * 86400000));
  if (typeof value !== 'string') return null;
  const match = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/.exec(value.trim());
  if (!match) return null;
  return new Date(Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0)));
}

function periodIndex(hour) {
  if (hour >= 9 && hour < 12) return 0;
  if (hour >= 12 && hour <= 16) return 1;
  if (hour >= 17 && hour <= 20) return 2;
  return -1;
}

function metric() { return { sum: 0, count: 0 }; }
function periodStats() { return { forecast: metric(), need: metric(), realNeed: metric() }; }
function valuePair() { return { firstSum: 0, secondSum: 0, count: 0 }; }

function addValues(target, first, second) {
  if (first == null || second == null) return;
  target.firstSum += first;
  target.secondSum += second;
  target.count += 1;
}

function valuePairAccuracy(value) {
  if (!value.count) return null;
  const max = Math.max(value.firstSum, value.secondSum);
  return max === 0 ? 1 : 1 - Math.abs(value.firstSum - value.secondSum) / max;
}

function addAccuracy(target, value) {
  if (value == null) return;
  target.sum += value;
  target.count += 1;
}

function addAggregatedPeriod(target, aggregate) {
  addAccuracy(target.forecast, valuePairAccuracy(aggregate.forecast));
  addAccuracy(target.need, valuePairAccuracy(aggregate.need));
  addAccuracy(target.realNeed, valuePairAccuracy(aggregate.realNeed));
}

function findColumn(headers, aliases) {
  const normalized = headers.map(normalize);
  for (const alias of aliases) {
    const index = normalized.indexOf(normalize(alias));
    if (index >= 0) return index;
  }
  return -1;
}

function cellValue(sheet, denseRows, rowIndex, columnIndex) {
  let cell;
  if (denseRows) {
    const row = denseRows[rowIndex];
    cell = row && row[columnIndex];
  } else {
    cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
  }
  return cell == null ? null : cell.v;
}

function analyze(buffer, sourceFile) {
  postMessage({ type: 'progress', message: 'Распаковка и чтение листа…' });
  const workbook = XLSX.read(buffer, {
    type: 'array',
    dense: true,
    cellDates: false,
    cellText: false,
    cellFormula: false,
    cellHTML: false,
    cellStyles: false,
    bookVBA: false,
    bookFiles: false
  });
  if (!workbook.SheetNames.length) throw new Error('в книге нет листов.');

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) throw new Error('в выбранном листе нет данных.');
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const denseRows = Array.isArray(sheet['!data']) ? sheet['!data'] : (Array.isArray(sheet) ? sheet : null);

  const headers = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    headers[column] = cellValue(sheet, denseRows, range.s.r, column);
  }
  const columns = {
    group: findColumn(headers, ['Стрит/ТЦ', 'Стрит ТЦ', 'Группа']),
    salon: findColumn(headers, ['Салон']),
    date: findColumn(headers, ['Дата']),
    historyTraffic: findColumn(headers, ['История трафик', 'Факт', 'История']),
    forecastTraffic: findColumn(headers, ['Прогноз трафик', 'План', 'Прогноз']),
    forecastNeed: findColumn(headers, ['Потребность от прогноза']),
    historyNeed: findColumn(headers, ['Потребность от истории']),
    realNeed: findColumn(headers, ['Потребность реальная', 'Потребность раеальая']),
    sales: findColumn(headers, ['Продажи'])
  };
  const required = [
    ['Стрит/ТЦ', columns.group], ['Салон', columns.salon], ['Дата', columns.date], ['История трафик', columns.historyTraffic],
    ['Прогноз трафик', columns.forecastTraffic], ['Потребность от прогноза', columns.forecastNeed],
    ['Потребность от истории', columns.historyNeed], ['Потребность реальная', columns.realNeed], ['Продажи', columns.sales]
  ];
  const missing = required.filter(function (item) { return item[1] < 0; }).map(function (item) { return item[0]; });
  if (missing.length) throw new Error('не найдены столбцы: ' + missing.join(', ') + '.');

  const data = {
    sourceFile: sourceFile,
    sheetName: sheetName,
    totalRows: Math.max(0, range.e.r - range.s.r),
    usedRows: 0,
    invalidRows: 0,
    outsideRows: 0,
    filteredLowAccuracyRows: 0,
    minDate: null,
    maxDate: null,
    periods: [periodStats(), periodStats(), periodStats()],
    salons: new Map()
  };
  const aggregates = new Map();
  const firstDataRow = range.s.r + 1;
  const totalRows = Math.max(0, range.e.r - range.s.r);

  for (let rowIndex = firstDataRow; rowIndex <= range.e.r; rowIndex += 1) {
    const salonValue = cellValue(sheet, denseRows, rowIndex, columns.salon);
    const groupValue = cellValue(sheet, denseRows, rowIndex, columns.group);
    const salonName = String(salonValue == null ? '' : salonValue).trim();
    const group = String(groupValue == null ? '' : groupValue).trim();
    const date = toDate(cellValue(sheet, denseRows, rowIndex, columns.date));
    if (!group || !salonName || !date || Number.isNaN(date.getTime())) {
      data.invalidRows += 1;
      if (denseRows) denseRows[rowIndex] = null;
      continue;
    }

    const period = periodIndex(date.getUTCHours());
    if (period < 0) {
      data.outsideRows += 1;
      if (denseRows) denseRows[rowIndex] = null;
      continue;
    }

    const historyTraffic = toNumber(cellValue(sheet, denseRows, rowIndex, columns.historyTraffic));
    const forecastTraffic = toNumber(cellValue(sheet, denseRows, rowIndex, columns.forecastTraffic));
    if (historyTraffic == null || forecastTraffic == null) {
      data.invalidRows += 1;
      if (denseRows) denseRows[rowIndex] = null;
      continue;
    }

    const forecastNeed = toNumber(cellValue(sheet, denseRows, rowIndex, columns.forecastNeed));
    const historyNeed = toNumber(cellValue(sheet, denseRows, rowIndex, columns.historyNeed));
    const realNeed = toNumber(cellValue(sheet, denseRows, rowIndex, columns.realNeed));
    const sales = toNumber(cellValue(sheet, denseRows, rowIndex, columns.sales));
    const day = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
    const aggregateKey = group + '\u001f' + salonName + '\u001f' + day + '\u001f' + period;
    let aggregate = aggregates.get(aggregateKey);
    if (!aggregate) {
      aggregate = {
        group: group,
        salon: salonName,
        day: day,
        period: period,
        forecast: valuePair(),
        need: valuePair(),
        realNeed: valuePair(),
        sales: 0,
        hasSales: false
      };
      aggregates.set(aggregateKey, aggregate);
    }
    addValues(aggregate.forecast, historyTraffic, forecastTraffic);
    addValues(aggregate.need, historyNeed, forecastNeed);
    addValues(aggregate.realNeed, realNeed, historyNeed);
    if (sales != null) { aggregate.sales += sales; aggregate.hasSales = true; }
    if (denseRows) denseRows[rowIndex] = null;

    const processed = rowIndex - range.s.r;
    if (processed % 25000 === 0 || processed === totalRows) {
      postMessage({ type: 'progress', message: 'Агрегация строк: ' + processed.toLocaleString('ru-RU') + ' из ' + totalRows.toLocaleString('ru-RU') });
    }
  }

  postMessage({ type: 'progress', message: 'Расчёт точности агрегированных записей…' });
  aggregates.forEach(function (aggregate) {
    const forecastAccuracy = valuePairAccuracy(aggregate.forecast);
    if (forecastAccuracy == null) { data.invalidRows += 1; return; }
    if (forecastAccuracy < FORECAST_ACCURACY_THRESHOLD) { data.filteredLowAccuracyRows += 1; return; }

    addAggregatedPeriod(data.periods[aggregate.period], aggregate);
    const salonKey = aggregate.group + '\u001f' + aggregate.salon;
    let salon = data.salons.get(salonKey);
    if (!salon) {
      salon = { group: aggregate.group, salon: aggregate.salon, dailySales: new Map(), periods: [periodStats(), periodStats(), periodStats()] };
      data.salons.set(salonKey, salon);
    }
    addAggregatedPeriod(salon.periods[aggregate.period], aggregate);
    if (aggregate.hasSales) salon.dailySales.set(aggregate.day, (salon.dailySales.get(aggregate.day) || 0) + aggregate.sales);
    data.usedRows += 1;
    if (data.minDate == null || aggregate.day < data.minDate) data.minDate = aggregate.day;
    if (data.maxDate == null || aggregate.day > data.maxDate) data.maxDate = aggregate.day;
  });
  aggregates.clear();
  return data;
}

self.onmessage = function (event) {
  if (!event.data || event.data.type !== 'analyze') return;
  try {
    const data = analyze(event.data.buffer, event.data.sourceFile || 'Excel');
    postMessage({ type: 'result', data: data });
  } catch (error) {
    postMessage({ type: 'error', message: error && error.message ? error.message : String(error) });
  }
};
