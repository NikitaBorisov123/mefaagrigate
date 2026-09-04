(function () {
  'use strict';

  const input = document.getElementById('hourly-file');
  const status = document.getElementById('hourly-status');
  const content = document.getElementById('hourly-report-content');
  if (!input || !status || !content) return;

  const periods = ['Утро', 'День', 'Вечер'];
  const salesBands = ['S_Более 150', 'S_100-150', 'S_50-100', 'S_Менее 50'];
  const pctFormat = new Intl.NumberFormat('ru-RU', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const numFormat = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });

  function esc(value) {
    return String(value == null ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
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
  function mergeMetric(target, source) {
    target.sum += source.sum;
    target.count += source.count;
  }
  function mergePeriod(target, source) {
    mergeMetric(target.forecast, source.forecast);
    mergeMetric(target.need, source.need);
    mergeMetric(target.realNeed, source.realNeed);
  }
  function findColumn(headers, aliases) {
    const normalized = headers.map(normalize);
    for (const alias of aliases) {
      const index = normalized.indexOf(normalize(alias));
      if (index >= 0) return index;
    }
    return -1;
  }

  function parseRows(rows, sourceFile, sheetName) {
    if (!rows.length) throw new Error('в выбранном листе нет данных.');
    const headers = rows[0] || [];
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
      totalRows: Math.max(0, rows.length - 1),
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

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const salonName = String(row[columns.salon] == null ? '' : row[columns.salon]).trim();
      const group = columns.group < 0 ? '' : String(row[columns.group] == null ? '' : row[columns.group]).trim();
      const date = toDate(row[columns.date]);
      if (!group || !salonName || !date || Number.isNaN(date.getTime())) { data.invalidRows += 1; continue; }
      const period = periodIndex(date.getUTCHours());
      if (period < 0) { data.outsideRows += 1; continue; }

      const historyTraffic = toNumber(row[columns.historyTraffic]);
      const forecastTraffic = toNumber(row[columns.forecastTraffic]);
      if (historyTraffic == null || forecastTraffic == null) { data.invalidRows += 1; continue; }
      const forecastNeed = toNumber(row[columns.forecastNeed]);
      const historyNeed = toNumber(row[columns.historyNeed]);
      const realNeed = toNumber(row[columns.realNeed]);
      const sales = toNumber(row[columns.sales]);
      const day = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
      const aggregateKey = group + '\u001f' + salonName + '\u001f' + day + '\u001f' + period;
      if (!aggregates.has(aggregateKey)) {
        aggregates.set(aggregateKey, {
          group: group,
          salon: salonName,
          day: day,
          period: period,
          forecast: valuePair(),
          need: valuePair(),
          realNeed: valuePair(),
          sales: 0,
          hasSales: false
        });
      }
      const aggregate = aggregates.get(aggregateKey);
      addValues(aggregate.forecast, historyTraffic, forecastTraffic);
      addValues(aggregate.need, historyNeed, forecastNeed);
      addValues(aggregate.realNeed, realNeed, historyNeed);
      if (sales != null) { aggregate.sales += sales; aggregate.hasSales = true; }
    }

    aggregates.forEach(function (aggregate) {
      const forecastAccuracy = valuePairAccuracy(aggregate.forecast);
      if (forecastAccuracy == null) { data.invalidRows += 1; return; }
      if (forecastAccuracy < 0.27) { data.filteredLowAccuracyRows += 1; return; }

      addAggregatedPeriod(data.periods[aggregate.period], aggregate);
      const salonKey = aggregate.group + '\u001f' + aggregate.salon;
      if (!data.salons.has(salonKey)) {
        data.salons.set(salonKey, { group: aggregate.group, salon: aggregate.salon, dailySales: new Map(), periods: [periodStats(), periodStats(), periodStats()] });
      }
      const salon = data.salons.get(salonKey);
      addAggregatedPeriod(salon.periods[aggregate.period], aggregate);
      if (aggregate.hasSales) salon.dailySales.set(aggregate.day, (salon.dailySales.get(aggregate.day) || 0) + aggregate.sales);
      data.usedRows += 1;
      if (data.minDate == null || aggregate.day < data.minDate) data.minDate = aggregate.day;
      if (data.maxDate == null || aggregate.day > data.maxDate) data.maxDate = aggregate.day;
    });
    return data;
  }

  function salesBand(salon) {
    const dailyValues = Array.from(salon.dailySales.values());
    const value = dailyValues.length ? dailyValues.reduce(function (sum, item) { return sum + item; }, 0) / dailyValues.length : 0;
    if (value > 150) return 'S_Более 150';
    if (value > 100) return 'S_100-150';
    if (value > 50) return 'S_50-100';
    return 'S_Менее 50';
  }
  function buildBandsForSalons(salons) {
    const result = {};
    salesBands.forEach(function (band) { result[band] = { salons: 0, periods: [periodStats(), periodStats(), periodStats()] }; });
    salons.forEach(function (salon) {
      const target = result[salesBand(salon)];
      target.salons += 1;
      for (let i = 0; i < 3; i += 1) mergePeriod(target.periods[i], salon.periods[i]);
    });
    return result;
  }
  function metricAccuracy(value) {
    return value.count ? value.sum / value.count : null;
  }
  function pct(value) {
    const accuracy = metricAccuracy(value);
    return accuracy == null ? '—' : pctFormat.format(accuracy);
  }
  function pctNumber(value) { return pctFormat.format(value); }
  function num(value) { return numFormat.format(value); }
  function metricRow(label, value) { return '<div class="metric"><span>' + esc(label) + '</span><b>' + pct(value) + '</b></div>'; }
  function metricFor(period, key) { return period[key]; }

  function bandTable(title, data, key) {
    const salons = Array.from(data.salons.values());
    const groups = Array.from(new Set(salons.map(function (salon) { return salon.group; }))).sort(function (a, b) { return a.localeCompare(b, 'ru'); });
    let rows = '';
    groups.forEach(function (group) {
      const groupBands = buildBandsForSalons(salons.filter(function (salon) { return salon.group.toLocaleLowerCase('ru-RU') === group.toLocaleLowerCase('ru-RU'); }));
      salesBands.forEach(function (band) {
        const value = groupBands[band];
        rows += '<tr><td><span class="tag">' + esc(group) + '</span></td><td>' + esc(band) + '</td><td>' + num(value.salons) + '</td><td>' + pct(metricFor(value.periods[0], key)) + '</td><td>' + pct(metricFor(value.periods[1], key)) + '</td><td>' + pct(metricFor(value.periods[2], key)) + '</td></tr>';
      });
    });
    return '<section class="panel"><h2>' + esc(title) + '</h2><div class="section-note">Группа продаж определяется по средним продажам каждого салона.</div><div class="scroll"><table><thead><tr><th>Группа</th><th>Группа продаж</th><th>Салонов</th><th>Утро</th><th>День</th><th>Вечер</th></tr></thead><tbody>' + rows + '</tbody></table></div></section>';
  }

  function periodAverage(data, key) {
    const available = data.periods.map(function (period) { return metricFor(period, key); }).filter(function (value) { return value.count > 0; });
    if (!available.length) return null;
    return available.reduce(function (sum, value) { return sum + metricAccuracy(value); }, 0) / available.length;
  }

  function summaryCard(label, value) {
    return '<div class="period-card"><div class="summary-value">' + (value == null ? '—' : pctNumber(value)) + '</div><div class="summary-label">' + esc(label) + '</div></div>';
  }

  function render(data) {
    const periodCards = periods.map(function (name, index) {
      return '<div class="period-card"><h3>' + name + '</h3>' + metricRow('Точность прогноза', data.periods[index].forecast) + metricRow('Точность потребности', data.periods[index].need) + metricRow('Точность потребности реальная', data.periods[index].realNeed) + '</div>';
    }).join('');
    const formatDate = function (time) { return new Date(time).toLocaleDateString('ru-RU', { timeZone: 'UTC' }); };
    const range = data.minDate == null ? 'даты не определены' : formatDate(data.minDate) + ' — ' + formatDate(data.maxDate);

    const summaryCards = summaryCard('Точность прогноза', periodAverage(data, 'forecast')) +
      summaryCard('Точность потребности', periodAverage(data, 'need')) +
      summaryCard('Точность потребности реальная', periodAverage(data, 'realNeed'));

    return '<section class="hero"><h1>Почасовая аналитика</h1><div class="sub">Источник: ' + esc(data.sourceFile) + ' · Лист: ' + esc(data.sheetName) + ' · Период: ' + esc(range) + ' · Салонов: ' + num(data.salons.size) + '</div></section>' +
      '<section class="panel"><h2>Среднее за три периода</h2><div class="section-note">Основа расчёта — агрегированные записи по каждому магазину, дате и периоду; фильтр 27% применяется после агрегации.</div><div class="period-grid">' + summaryCards + '</div></section>' +
      '<section class="panel"><h2>Общая точность по периодам дня</h2><div class="period-grid">' + periodCards + '</div></section>' +
      bandTable('Точность прогноза по группам продаж', data, 'forecast') +
      bandTable('Точность потребности по группам продаж', data, 'need') +
      bandTable('Точность реальной потребности по группам продаж', data, 'realNeed');
  }

  input.addEventListener('change', async function () {
    const file = input.files && input.files[0];
    if (!file) return;
    if (!window.XLSX) { status.className = 'status error'; status.textContent = 'Не удалось загрузить модуль чтения Excel.'; return; }
    input.disabled = true;
    status.className = 'status';
    status.textContent = 'Чтение и анализ файла…';
    await new Promise(function (resolve) { setTimeout(resolve, 30); });
    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: false, cellText: false });
      if (!workbook.SheetNames.length) throw new Error('в книге нет листов.');
      const sheetName = workbook.SheetNames[0];
      const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
      const data = parseRows(rows, file.name, sheetName);
      if (!data.usedRows) throw new Error('нет агрегированных записей в интервале 09:00–20:59 после применения порога 27%.');
      content.innerHTML = render(data);
      status.className = 'status ok';
      status.textContent = 'Загружен «' + file.name + '»: ' + num(data.usedRows) + ' агрегированных записей включено, ' + num(data.filteredLowAccuracyRows) + ' исключено по порогу 27%.';
    } catch (error) {
      status.className = 'status error';
      status.textContent = 'Ошибка: ' + (error && error.message ? error.message : String(error));
    } finally {
      input.disabled = false;
    }
  });
})();
