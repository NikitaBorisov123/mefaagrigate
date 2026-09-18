(function () {
  'use strict';

  const input = document.getElementById('excel-file');
  const status = document.getElementById('upload-status');
  const content = document.getElementById('report-content');
  if (!input || !status || !content) return;

  const numberFormat = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
  const decimalFormat = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const percentFormat = new Intl.NumberFormat('ru-RU', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function normalizeHeader(value) {
    return String(value == null ? '' : value)
      .trim()
      .toLocaleLowerCase('ru-RU')
      .replace(/ё/g, 'е')
      .replace(/\s+/g, ' ');
  }

  function toNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string') return null;
    const normalized = value.replace(/\s/g, '').replace(',', '.');
    if (!normalized) return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function toDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      const date = new Date(Math.round((value - 25569) * 86400000));
      return Number.isNaN(date.getTime()) ? null : date;
    }
    if (typeof value === 'string') {
      const text = value.trim();
      let match = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/.exec(text);
      if (match) return new Date(Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1])));
      match = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
      if (match) return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
    }
    return null;
  }

  function newStats() {
    return {
      rows: 0,
      zeroZeroRows: 0,
      sumPlan: 0,
      sumActual: 0,
      sumMax: 0,
      sumAbsDiff: 0,
      sumAccuracy: 0,
      sumNeedAccuracy: 0,
      needAccuracyCount: 0,
      sumRealNeedAccuracy: 0,
      realNeedAccuracyCount: 0,
      metrics: [
        { sum: 0, count: 0 },
        { sum: 0, count: 0 },
        { sum: 0, count: 0 },
        { sum: 0, count: 0 },
        { sum: 0, count: 0 },
        { sum: 0, count: 0 }
      ]
    };
  }

  function addStats(stats, plan, actual, metrics) {
    const max = Math.max(plan, actual);
    const diff = Math.abs(plan - actual);
    let accuracy;
    if (max === 0) {
      accuracy = 1;
      stats.zeroZeroRows += 1;
    } else {
      accuracy = 1 - diff / max;
    }
    stats.rows += 1;
    stats.sumPlan += plan;
    stats.sumActual += actual;
    stats.sumMax += max;
    stats.sumAbsDiff += diff;
    stats.sumAccuracy += accuracy;
    if (metrics[0] != null && metrics[1] != null) {
      const needMax = Math.max(metrics[0], metrics[1]);
      const needDiff = Math.abs(metrics[0] - metrics[1]);
      stats.sumNeedAccuracy += needMax === 0 ? 1 : 1 - needDiff / needMax;
      stats.needAccuracyCount += 1;
    }
    if (metrics[0] != null && metrics[2] != null) {
      const realNeedMax = Math.max(metrics[0], metrics[2]);
      const realNeedDiff = Math.abs(metrics[0] - metrics[2]);
      stats.sumRealNeedAccuracy += realNeedMax === 0 ? 1 : 1 - realNeedDiff / realNeedMax;
      stats.realNeedAccuracyCount += 1;
    }
    metrics.forEach(function (value, index) {
      if (value == null) return;
      stats.metrics[index].sum += value;
      stats.metrics[index].count += 1;
    });
  }

  function calculated(stats) {
    const totalsMax = Math.max(stats.sumPlan, stats.sumActual);
    return {
      averageAccuracy: stats.rows ? stats.sumAccuracy / stats.rows : 0,
      needAccuracy: stats.needAccuracyCount ? stats.sumNeedAccuracy / stats.needAccuracyCount : 0,
      realNeedAccuracy: stats.realNeedAccuracyCount ? stats.sumRealNeedAccuracy / stats.realNeedAccuracyCount : null,
      weightedAccuracy: stats.sumMax ? 1 - stats.sumAbsDiff / stats.sumMax : 1,
      totalsAccuracy: totalsMax ? 1 - Math.abs(stats.sumPlan - stats.sumActual) / totalsMax : 1,
      planExecution: stats.sumPlan ? stats.sumActual / stats.sumPlan : 0
    };
  }

  function findColumn(header, aliases) {
    const normalized = header.map(normalizeHeader);
    for (const alias of aliases) {
      const index = normalized.indexOf(normalizeHeader(alias));
      if (index >= 0) return index;
    }
    return -1;
  }

  function parseRows(rows, sourceFile, sheetName) {
    if (!rows.length) throw new Error('В выбранном листе нет данных.');
    const header = rows[0];
    const columns = {
      group: findColumn(header, ['Стрит/ТЦ', 'Стрит ТЦ', 'Группа']),
      salon: findColumn(header, ['Салон']),
      date: findColumn(header, ['Дата']),
      actual: findColumn(header, ['Факт']),
      plan: findColumn(header, ['План']),
      historyNeed: findColumn(header, ['Потребность от истории']),
      forecastNeed: findColumn(header, ['Потребность от прогноза']),
      realNeed: findColumn(header, ['Потребность реальная', 'Потребность раеальая']),
      sales: findColumn(header, ['Продажи']),
      coverage: findColumn(header, ['Покрытие']),
      utilization: findColumn(header, ['Утилизация'])
    };
    const required = [['Стрит/ТЦ', columns.group], ['Салон', columns.salon], ['Факт', columns.actual], ['План', columns.plan]];
    const missing = required.filter(function (item) { return item[1] < 0; }).map(function (item) { return item[0]; });
    if (missing.length) throw new Error('Не найдены обязательные столбцы: ' + missing.join(', ') + '.');

    const data = {
      sourceFile: sourceFile,
      sheetName: sheetName,
      totalRows: Math.max(0, rows.length - 1),
      invalidRows: 0,
      filteredLowAccuracyRows: 0,
      minDate: null,
      maxDate: null,
      overall: newStats(),
      groups: new Map(),
      salons: new Map()
    };

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const group = String(row[columns.group] == null ? '' : row[columns.group]).trim();
      const salon = String(row[columns.salon] == null ? '' : row[columns.salon]).trim();
      const plan = toNumber(row[columns.plan]);
      const actual = toNumber(row[columns.actual]);
      if (!group || !salon || plan == null || actual == null) {
        data.invalidRows += 1;
        continue;
      }

      const optionalIndexes = [columns.historyNeed, columns.forecastNeed, columns.realNeed, columns.sales, columns.coverage, columns.utilization];
      const metrics = optionalIndexes.map(function (index) { return index < 0 ? null : toNumber(row[index]); });
      const forecastMax = Math.max(plan, actual);
      const dailyForecastAccuracy = forecastMax === 0 ? 1 : 1 - Math.abs(plan - actual) / forecastMax;
      if (dailyForecastAccuracy < 0.20) {
        data.filteredLowAccuracyRows += 1;
        continue;
      }
      addStats(data.overall, plan, actual, metrics);

      if (!data.groups.has(group)) data.groups.set(group, newStats());
      addStats(data.groups.get(group), plan, actual, metrics);

      const salonKey = group + '\u001f' + salon;
      if (!data.salons.has(salonKey)) data.salons.set(salonKey, { group: group, salon: salon, values: newStats() });
      addStats(data.salons.get(salonKey).values, plan, actual, metrics);

      if (columns.date >= 0) {
        const date = toDate(row[columns.date]);
        if (date) {
          if (!data.minDate || date < data.minDate) data.minDate = date;
          if (!data.maxDate || date > data.maxDate) data.maxDate = date;
        }
      }
    }
    return data;
  }

  function pct(value) { return percentFormat.format(value); }
  function pctOrDash(value) { return value == null ? '—' : pct(value); }
  function num(value) { return numberFormat.format(value); }
  function dec(value) { return decimalFormat.format(value); }
  function avg(metric) { return metric.count ? dec(metric.sum / metric.count) : '—'; }
  function accuracyClass(value) { return value >= 0.85 ? 'good' : value < 0.75 ? 'warn' : ''; }
  function shorten(value, max) { return value.length <= max ? value : value.slice(0, max - 1) + '…'; }

  function formatDate(date) {
    return date ? date.toLocaleDateString('ru-RU', { timeZone: 'UTC' }) : null;
  }

  function salonTable(title, rows) {
    return '<div class="panel"><h2>' + esc(title) + '</h2><div class="scroll"><table><thead><tr><th>Салон</th><th>Группа</th><th>Дней</th><th>Среднедневная точность</th></tr></thead><tbody>' +
      rows.map(function (item) {
        return '<tr><td title="' + esc(item.salon) + '">' + esc(shorten(item.salon, 34)) + '</td><td>' + esc(item.group) + '</td><td>' + num(item.values.rows) + '</td><td>' + pct(calculated(item.values).averageAccuracy) + '</td></tr>';
      }).join('') + '</tbody></table></div></div>';
  }

  function selectedSalons(salonEntries) {
    const requested = [
      'G470, СЗ_Колпино_Пролетарская_36_МеркурийТЦ',
      'V339, ПВ_Уфа_Первомайская_98',
      'U551, УР_Нефтеюганск_Усть-Балыкская_ТЦ_секц',
      'U408, УР_Челябинск_Комарова/Салютная_2',
      'A084, СТ_Балашиха_Энтузиастов_54А',
      'D083, ДВ_Чита_Бабушкина_64',
      'K984, КВ_Липецк_Краснозаводская_23_ТЦ',
      'D164, ДВ_Улан-Удэ_Автомобилистов_4А_МоллТЦ',
      'K962, КВ_Сочи_Ленина_1А_ТатулянТЦ',
      'C022, ЦР_Дзержинск_Гайдара_61Г_645',
      'A72N, СТ_Звенигород_Московская_20_15',
      'C240, ЦР_Брянск_2яМичурина_42_МБТЦ',
      'S576, СБ_Кызыл_ТувинскихДобровольцев_24',
      'G091, СЗ_Псков_Ленина_7А',
      'A67N, СТ_Павловский Посад_Привокзальная_19'
    ];
    const rows = requested.map(function (requestedSalon) {
      const code = requestedSalon.split(',')[0].trim().toLocaleLowerCase('ru-RU');
      const salon = salonEntries.find(function (item) { return item.salon.toLocaleLowerCase('ru-RU').startsWith(code + ','); });
      return '<tr><td>' + esc(salon ? salon.salon : requestedSalon) + '</td><td>' + (salon ? esc(salon.group) : '—') + '</td><td>' + (salon ? pct(calculated(salon.values).averageAccuracy) : 'Нет данных') + '</td></tr>';
    }).join('');
    return '<section class="panel selected-salons"><h2>Точность прогноза выбранных магазинов</h2><div class="section-note">Средняя точность за период; дни с точностью ниже 20% исключены.</div><div class="scroll"><table><thead><tr><th>Магазин</th><th>Группа</th><th>Средняя точность</th></tr></thead><tbody>' + rows + '</tbody></table></div></section>';
  }

  function salesBand(salon) {
    const metric = salon.values.metrics[3];
    const value = metric.count ? metric.sum / metric.count : 0;
    if (value > 150) return 'S_Более 150';
    if (value > 100) return 'S_100-150';
    if (value > 50) return 'S_50-100';
    return 'S_Менее 50';
  }

  function accuracyBand(value) {
    if (value > 0.95) return 'Т_95-100';
    if (value > 0.90) return 'Т_90-95';
    if (value > 0.85) return 'Т_85-90';
    if (value > 0.80) return 'Т_80-85';
    if (value > 0.70) return 'Т_70-80';
    if (value > 0.60) return 'Т_60-70';
    return 'Т_0-60';
  }

  function salesPivot(groupEntries, salonEntries) {
    const bands = ['S_Более 150', 'S_100-150', 'S_50-100', 'S_Менее 50'];
    let header = '<th>Группа продаж</th>';
    groupEntries.forEach(function (item) { header += '<th>' + esc(item[0]) + '</th>'; });
    header += '<th>Итого</th><th>Среднедневная точность</th>';
    const rows = bands.map(function (band) {
      const bandSalons = salonEntries.filter(function (salon) { return salesBand(salon) === band; });
      let total = 0;
      let cells = '';
      groupEntries.forEach(function (item) {
        const group = item[0];
        const count = salonEntries.filter(function (salon) {
          return salon.group.toLocaleLowerCase('ru-RU') === group.toLocaleLowerCase('ru-RU') && salesBand(salon) === band;
        }).length;
        total += count;
        cells += '<td>' + num(count) + '</td>';
      });
      const accuracyRows = bandSalons.reduce(function (sum, salon) { return sum + salon.values.rows; }, 0);
      const accuracySum = bandSalons.reduce(function (sum, salon) { return sum + salon.values.sumAccuracy; }, 0);
      const accuracy = accuracyRows ? accuracySum / accuracyRows : 0;
      return '<tr><td>' + esc(band) + '</td>' + cells + '<td><b>' + num(total) + '</b></td><td class="' + accuracyClass(accuracy) + '">' + pct(accuracy) + '</td></tr>';
    }).join('');
    return '<section class="panel"><h2>Свод по группам продаж</h2><div class="section-note">Группа продаж определяется по среднему значению столбца «Продажи» для каждого салона.</div><div class="scroll"><table><thead><tr>' + header + '</tr></thead><tbody>' + rows + '</tbody></table></div></section>';
  }

  function operationalBySales(groupEntries, salonEntries) {
    const bands = ['S_Более 150', 'S_100-150', 'S_50-100', 'S_Менее 50'];
    let rows = '';
    groupEntries.forEach(function (groupItem) {
      const group = groupItem[0];
      bands.forEach(function (band) {
        const salons = salonEntries.filter(function (salon) {
          return salon.group.toLocaleLowerCase('ru-RU') === group.toLocaleLowerCase('ru-RU') && salesBand(salon) === band;
        });
        const coverage = salons.reduce(function (result, salon) {
          result.sum += salon.values.metrics[4].sum;
          result.count += salon.values.metrics[4].count;
          return result;
        }, { sum: 0, count: 0 });
        const utilization = salons.reduce(function (result, salon) {
          result.sum += salon.values.metrics[5].sum;
          result.count += salon.values.metrics[5].count;
          return result;
        }, { sum: 0, count: 0 });
        rows += '<tr><td><span class="tag">' + esc(group) + '</span></td><td>' + esc(band) + '</td><td>' + num(salons.length) + '</td><td>' + avg(coverage) + '</td><td>' + avg(utilization) + '</td></tr>';
      });
    });
    return '<section class="panel"><h2>Операционные показатели</h2><div class="section-note">Покрытие и утилизация показаны по группам СТРИТ/ТЦ и группам продаж.</div><div class="scroll"><table><thead><tr><th>Группа</th><th>Группа продаж</th><th>Салонов</th><th>Покрытие</th><th>Утилизация</th></tr></thead><tbody>' + rows + '</tbody></table></div></section>';
  }

  function accuracyDistribution(groupEntries, salonEntries) {
    const bands = ['Т_95-100', 'Т_90-95', 'Т_85-90', 'Т_80-85', 'Т_70-80', 'Т_60-70', 'Т_0-60'];
    const totals = {};
    bands.forEach(function (band) {
      totals[band] = salonEntries.filter(function (salon) { return accuracyBand(calculated(salon.values).averageAccuracy) === band; }).length;
    });
    const maxTotal = Math.max(1, Math.max.apply(null, bands.map(function (band) { return totals[band]; })));
    const legend = groupEntries.map(function (item, index) {
      return '<span><i class="dot" style="background:' + (index % 2 === 0 ? '#176b73' : '#ef8354') + '"></i>' + esc(item[0]) + '</span>';
    }).join('');
    const rows = bands.map(function (band) {
      let segments = '';
      groupEntries.forEach(function (item, index) {
        const group = item[0];
        const count = salonEntries.filter(function (salon) {
          return salon.group.toLocaleLowerCase('ru-RU') === group.toLocaleLowerCase('ru-RU') && accuracyBand(calculated(salon.values).averageAccuracy) === band;
        }).length;
        if (!count) return;
        const width = count * 100 / maxTotal;
        segments += '<span class="dist-seg g' + (index % 2) + '" style="width:' + width.toFixed(2) + '%" title="' + esc(group) + ': ' + num(count) + '">' + num(count) + '</span>';
      });
      return '<div class="dist-row"><div class="dist-label">' + esc(band) + '</div><div class="dist-track">' + segments + '</div><div class="dist-total">' + num(totals[band]) + '</div></div>';
    }).join('');
    return '<section class="panel"><h2>Распределение салонов по точности прогноза</h2><div class="section-note">Для каждого салона рассчитана средняя точность по всем оставшимся дням.</div><div class="legend">' + legend + '</div><div class="dist">' + rows + '</div></section>';
  }

  function accuracyShares(salonEntries) {
    const total = salonEntries.length;
    const atLeast80 = salonEntries.filter(function (salon) { return calculated(salon.values).averageAccuracy >= 0.80; }).length;
    const atLeast60 = salonEntries.filter(function (salon) { return calculated(salon.values).averageAccuracy >= 0.60; }).length;
    const cards = [
      ['80–100%', atLeast80],
      ['60–100%', atLeast60],
      ['Ниже 60%', total - atLeast60]
    ].map(function (item) {
      return '<div class="accuracy-share-card"><div class="accuracy-share-value">' + (total ? pct(item[1] / total) : '—') + '</div><div class="accuracy-share-label">' + item[0] + '</div><div class="accuracy-share-count">' + num(item[1]) + ' из ' + num(total) + ' магазинов</div></div>';
    }).join('');
    return '<section class="panel accuracy-shares"><h2>Доля магазинов по точности прогноза</h2><div class="section-note">Средняя точность каждого магазина рассчитана по оставшимся дням. Группы пересекаются: 60–100% включает 80–100%.</div><div class="accuracy-share-grid">' + cards + '</div></section>';
  }

  function accuracySalesMatrix(groupEntries, salonEntries) {
    const salesBands = ['S_Менее 50', 'S_50-100', 'S_100-150', 'S_Более 150'];
    const accuracyBands = ['Т_0-60', 'Т_60-70', 'Т_70-80', 'Т_80-85', 'Т_85-90', 'Т_90-95', 'Т_95-100'];
    const sections = groupEntries.map(function (groupItem) {
      const group = groupItem[0];
      const counts = accuracyBands.map(function () { return salesBands.map(function () { return 0; }); });
      const salesTotals = salesBands.map(function () { return 0; });
      const accuracyTotals = accuracyBands.map(function () { return 0; });
      let groupTotal = 0;
      salonEntries.forEach(function (salon) {
        if (salon.group.toLocaleLowerCase('ru-RU') !== group.toLocaleLowerCase('ru-RU')) return;
        const accuracyIndex = accuracyBands.indexOf(accuracyBand(calculated(salon.values).averageAccuracy));
        const salesIndex = salesBands.indexOf(salesBand(salon));
        if (accuracyIndex < 0 || salesIndex < 0) return;
        counts[accuracyIndex][salesIndex] += 1;
        salesTotals[salesIndex] += 1;
        accuracyTotals[accuracyIndex] += 1;
        groupTotal += 1;
      });

      const maxShare = groupTotal ? Math.max.apply(null, counts.flat().map(function (count) { return count * 100 / groupTotal; })) : 0;
      const axisMax = Math.max(5, Math.ceil(maxShare / 5) * 5);
      const axis = [4, 3, 2, 1, 0].map(function (tick) { return '<span>' + (axisMax * tick / 4).toLocaleString('ru-RU', { maximumFractionDigits: 2 }) + '%</span>'; }).join('');
      const bars = salesBands.map(function (salesBandName, salesIndex) {
        return '<div class="matrix-bars">' + accuracyBands.map(function (accuracyBandName, accuracyIndex) {
          const count = counts[accuracyIndex][salesIndex];
          if (!count) return '';
          const share = count * 100 / groupTotal;
          const height = (share * 100 / axisMax).toFixed(2);
          const label = share >= 0.8 ? '<span class="matrix-bar-label">' + num(Math.round(share)) + '%</span>' : '';
          const title = group + ' · ' + salesBandName + ' · ' + accuracyBandName + ': ' + num(count) + ' магазинов (' + share.toLocaleString('ru-RU', { maximumFractionDigits: 2 }) + '%)';
          return '<span class="matrix-bar acc-' + accuracyIndex + '" style="height:' + height + '%" title="' + esc(title) + '">' + label + '</span>';
        }).join('') + '</div>';
      }).join('');
      const legend = accuracyBands.map(function (band, index) { return '<span><i class="acc-' + index + '"></i>' + esc(band) + '</span>'; }).join('');
      const header = salesBands.map(function (band) { return '<th>' + esc(band) + '</th>'; }).join('');
      const totals = salesTotals.map(function (count) { return '<td>' + num(count) + '</td>'; }).join('');
      const rows = accuracyBands.map(function (band, accuracyIndex) {
        return '<tr><td>' + esc(band) + '</td>' + counts[accuracyIndex].map(function (count) { return '<td>' + num(count) + '</td>'; }).join('') + '<td>' + num(accuracyTotals[accuracyIndex]) + '</td></tr>';
      }).join('');

      return '<div class="matrix-section"><h3>' + esc(group) + ' · ' + num(groupTotal) + ' магазинов</h3>' +
        '<div class="matrix-visual"><div class="matrix-inner"><div class="matrix-chart"><div class="matrix-axis">' + axis + '</div><div class="matrix-plot">' + bars + '</div></div>' +
        '<div class="matrix-x-axis">' + salesBands.map(function (band) { return '<span>' + esc(band) + '</span>'; }).join('') + '</div><div class="matrix-legend">' + legend + '</div></div></div>' +
        '<div class="scroll"><table class="matrix-table"><thead><tr><th>Группа точности</th>' + header + '<th>Итого</th></tr></thead><tbody>' +
        '<tr class="matrix-total"><td>Итого</td>' + totals + '<td class="matrix-grand-total">' + num(groupTotal) + '</td></tr>' + rows + '</tbody></table></div></div>';
    }).join('');
    return '<section class="panel"><h2>Точность прогноза по группам продаж</h2><div class="section-note">Каждый магазин учитывается один раз. Проценты на графике — доля от всех магазинов соответствующего типа; в матрице показано количество магазинов.</div>' + sections + '</section>';
  }

  function renderReport(data) {
    const overall = calculated(data.overall);
    const groupEntries = Array.from(data.groups.entries()).sort(function (a, b) { return a[0].localeCompare(b[0], 'ru'); });
    const salonEntries = Array.from(data.salons.values());
    const range = data.minDate && data.maxDate ? formatDate(data.minDate) + ' — ' + formatDate(data.maxDate) : 'даты не определены';
    const now = new Date().toLocaleString('ru-RU', { dateStyle: 'short', timeStyle: 'short' });

    const groupRows = groupEntries.map(function (item) {
      const key = item[0], stats = item[1], calc = calculated(stats);
      const salons = salonEntries.filter(function (salon) { return salon.group.toLocaleLowerCase('ru-RU') === key.toLocaleLowerCase('ru-RU'); }).length;
      return '<tr><td><span class="tag">' + esc(key) + '</span></td><td>' + num(salons) + '</td><td>' + pct(calc.planExecution) + '</td><td class="' + accuracyClass(calc.averageAccuracy) + '">' + pct(calc.averageAccuracy) + '</td><td>' + pct(calc.needAccuracy) + '</td><td>' + pctOrDash(calc.realNeedAccuracy) + '</td></tr>';
    }).join('');

    const leaders = salonEntries.slice().sort(function (a, b) { return calculated(b.values).averageAccuracy - calculated(a.values).averageAccuracy || b.values.rows - a.values.rows; }).slice(0, 10);
    const attention = salonEntries.slice().sort(function (a, b) { return calculated(a.values).averageAccuracy - calculated(b.values).averageAccuracy || b.values.rows - a.values.rows; }).slice(0, 10);

    return '<section class="hero"><h1>Аналитический отчёт по прогнозу</h1><div class="sub">Источник: ' + esc(data.sourceFile) + ' · Лист: ' + esc(data.sheetName) + ' · Период: ' + esc(range) + ' · Сформирован: ' + esc(now) + '</div></section>' +
      '<section class="grid kpis"><div class="card"><div class="value">' + pct(overall.averageAccuracy) + '</div><div class="label">Среднедневная точность</div></div><div class="card"><div class="value">' + pct(overall.needAccuracy) + '</div><div class="label">Точность потребности</div></div><div class="card"><div class="value">' + pctOrDash(overall.realNeedAccuracy) + '</div><div class="label">Точность потребности реальная</div></div><div class="card"><div class="value">' + pct(overall.planExecution) + '</div><div class="label">Факт / план</div></div><div class="card"><div class="value">' + num(data.salons.size) + '</div><div class="label">Салонов</div></div><div class="card"><div class="value">' + avg(data.overall.metrics[4]) + '</div><div class="label">Среднее покрытие</div></div><div class="card"><div class="value">' + avg(data.overall.metrics[5]) + '</div><div class="label">Средняя утилизация</div></div></section>' +
      accuracyShares(salonEntries) +
      '<section class="panel"><h2>Сравнение СТРИТ и ТЦ</h2><div class="section-note">Среднедневная точность рассчитана как среднее значений точности по всем оставшимся дням.</div><div class="scroll"><table><thead><tr><th>Группа</th><th>Салонов</th><th>Факт / план</th><th>Среднедневная точность</th><th>Точность потребности</th><th>Точность потребности реальная</th></tr></thead><tbody>' + groupRows + '</tbody></table></div></section>' +
      salesPivot(groupEntries, salonEntries) +
      operationalBySales(groupEntries, salonEntries) +
      accuracyDistribution(groupEntries, salonEntries) +
      accuracySalesMatrix(groupEntries, salonEntries) +
      '<section class="grid two">' + salonTable('Лидеры по точности', leaders) + salonTable('Зоны внимания', attention) + '</section>' +
      selectedSalons(salonEntries);
  }

  input.addEventListener('change', async function () {
    const file = input.files && input.files[0];
    if (!file) return;
    if (!window.XLSX) {
      status.className = 'upload-status error';
      status.textContent = 'Не удалось загрузить модуль чтения Excel. Проверьте файл xlsx.full.min.js.';
      return;
    }

    input.disabled = true;
    status.className = 'upload-status';
    status.textContent = 'Чтение и анализ файла…';
    await new Promise(function (resolve) { setTimeout(resolve, 30); });

    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: false, cellText: false });
      if (!workbook.SheetNames.length) throw new Error('В книге нет листов.');
      const sheetName = workbook.SheetNames[0];
      const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
      const data = parseRows(rows, file.name, sheetName);
      if (!data.overall.rows) throw new Error('Не найдено строк с заполненными группой, салоном, планом и фактом.');
      content.innerHTML = renderReport(data);
      status.className = 'upload-status ok';
      status.textContent = 'Загружен «' + file.name + '»: ' + num(data.overall.rows) + ' строк включено, ' + num(data.filteredLowAccuracyRows) + ' исключено по порогу 20%.';
    } catch (error) {
      status.className = 'upload-status error';
      status.textContent = 'Ошибка: ' + (error && error.message ? error.message : String(error));
    } finally {
      input.disabled = false;
    }
  });
})();
