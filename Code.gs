/**
 * Dashboard Avaliação MEC — UNISUAM
 * Backend Google Apps Script (dados em JSON no Drive)
 */

var CONFIG = {
  FOLDER_NAME: 'Dashboard Avaliacao MEC - Dados',
  DATA_FILE_NAME: 'cursos_data.json'
};

var INDICADORES_LISTA = (function () {
  var list = [];
  var i;
  for (i = 1; i <= 24; i++) list.push('1.' + i);
  for (i = 1; i <= 16; i++) list.push('2.' + i);
  for (i = 1; i <= 18; i++) list.push('3.' + i);
  return list;
})();

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Avaliação MEC | UNISUAM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Data URL do logo UNISUAM, lida de LogoData.html. */
function getLogoDataUrl() {
  var html = HtmlService.createHtmlOutputFromFile('LogoData').getContent();
  var match = String(html).match(/LOGO_UNISUAM_B64\s*=\s*"([^"]+)"/);
  if (!match) return '';
  return 'data:image/png;base64,' + match[1];
}

/** Garante pasta + arquivo JSON no Drive do usuário. */
function getOrCreateDataFile_() {
  var folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);

  var files = folder.getFilesByName(CONFIG.DATA_FILE_NAME);
  if (files.hasNext()) {
    return files.next();
  }

  var seed = typeof SEED_DATA !== 'undefined' ? SEED_DATA : { version: 1, cursos: [] };
  seed.updatedAt = new Date().toISOString();
  var file = folder.createFile(
    CONFIG.DATA_FILE_NAME,
    JSON.stringify(seed, null, 2),
    MimeType.PLAIN_TEXT
  );
  return file;
}

function readData_() {
  var file = getOrCreateDataFile_();
  var raw = file.getBlob().getDataAsString('UTF-8');
  var data = JSON.parse(raw || '{"version":1,"cursos":[]}');
  if (!data.cursos) data.cursos = [];
  var beforeSig = cursosSignature_(data.cursos);
  var before = data.cursos.length;
  data.cursos = dedupeCursos_(data.cursos);
  var removed = before - data.cursos.length;
  var normalized = cursosSignature_(data.cursos) !== beforeSig;
  if (!data.skipSeedMerge) {
    data = mergeMissingScoresFromSeed_(data);
  }
  var beforeSparse = data.cursos.length;
  data.cursos = dropSparseCursos_(data.cursos);
  var removedSparse = beforeSparse - data.cursos.length;
  if (removed > 0 || removedSparse > 0 || normalized) writeData_(data);
  data.removedDuplicates = removed;
  data.removedSparse = removedSparse;
  return data;
}

/**
 * Preenche conceitos/indicadores vazios a partir do SEED_DATA.
 * Não sobrescreve valores já existentes (importados pelo usuário).
 */
function mergeMissingScoresFromSeed_(data) {
  if (typeof SEED_DATA === 'undefined' || !SEED_DATA.cursos || !SEED_DATA.cursos.length) {
    return data;
  }

  var seedByKey = {};
  SEED_DATA.cursos.forEach(function (s) {
    seedByKey[courseKey_(s)] = s;
  });

  var changed = false;
  data.cursos.forEach(function (c, idx) {
    var s = seedByKey[courseKey_(c)];
    if (!s) {
      for (var i = 0; i < SEED_DATA.cursos.length; i++) {
        var cand = SEED_DATA.cursos[i];
        if (String(cand.codigo) === String(c.codigo) &&
            Number(cand.anoVisita) === Number(c.anoVisita) &&
            String(cand.unidade || '') === String(c.unidade || '')) {
          s = cand;
          break;
        }
      }
    }
    if (!s) return;

    if ((c.conceitoFinalContinuo == null || c.conceitoFinalContinuo === '') &&
        s.conceitoFinalContinuo != null) {
      c.conceitoFinalContinuo = s.conceitoFinalContinuo;
      changed = true;
    }
    if ((c.conceitoFinalFaixa == null || c.conceitoFinalFaixa === '') &&
        s.conceitoFinalFaixa != null) {
      c.conceitoFinalFaixa = s.conceitoFinalFaixa;
      changed = true;
    }
    c.indicadores = c.indicadores || {};
    if (s.indicadores) {
      Object.keys(s.indicadores).forEach(function (k) {
        if ((c.indicadores[k] == null || c.indicadores[k] === '') && s.indicadores[k] != null) {
          c.indicadores[k] = s.indicadores[k];
          changed = true;
        }
      });
    }
    data.cursos[idx] = c;
  });

  if (changed) writeData_(data);
  return data;
}

function writeData_(data) {
  data.updatedAt = new Date().toISOString();
  data.version = data.version || 1;
  var file = getOrCreateDataFile_();
  file.setContent(JSON.stringify(data, null, 2));
  return data;
}

function getDashboardData() {
  try {
    var data = readData_();
    return {
      ok: true,
      data: data,
      removedDuplicates: data.removedDuplicates || 0,
      removedSparse: data.removedSparse || 0,
      folderUrl: getDataFolderUrl_(),
      indicadores: INDICADORES_LISTA
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function getDataFolderUrl_() {
  var folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) return folders.next().getUrl();
  return null;
}

/**
 * Substitui ou faz merge de cursos importados.
 * strategy: 'replace_all' | 'upsert'
 */
function saveCursos(cursos, strategy) {
  try {
    if (!Array.isArray(cursos)) {
      return { ok: false, error: 'Payload inválido.' };
    }
    strategy = strategy || 'upsert';
    var data = readData_();

    var incomingCount = 0;
    if (strategy === 'replace_all') {
      var normalized = [];
      cursos.forEach(function (raw) {
        var c = normalizeCurso_(raw);
        var key = courseKey_(c);
        if (!key || key === '||') return;
        normalized.push(c);
        incomingCount++;
      });
      data.cursos = dedupeCursos_(normalized);
    } else {
      cursos.forEach(function (incoming) {
        var c = normalizeCurso_(incoming);
        var key = courseKey_(c);
        if (!key || key === '||') return;
        incomingCount++;
        var idx = -1;
        for (var i = 0; i < data.cursos.length; i++) {
          if (courseKey_(data.cursos[i]) === key) {
            idx = i;
            break;
          }
        }
        if (idx >= 0) {
          data.cursos[idx] = mergeCursoPreserveScores_(data.cursos[idx], c);
        } else {
          data.cursos.push(c);
        }
      });
      var beforeDedupe = data.cursos.length;
      data.cursos = dedupeCursos_(data.cursos);
      incomingCount = beforeDedupe;
    }

    var beforeSparse = data.cursos.length;
    data.cursos = dropSparseCursos_(data.cursos);
    var removedSparse = beforeSparse - data.cursos.length;
    var removedDuplicates = Math.max(0, incomingCount - data.cursos.length - removedSparse);
    data.skipSeedMerge = true;
    delete data.removedDuplicates;
    delete data.removedSparse;
    writeData_(data);
    return {
      ok: true,
      total: data.cursos.length,
      removedDuplicates: removedDuplicates,
      removedSparse: removedSparse,
      data: data
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function saveCursoManual(curso) {
  return saveCursos([curso], 'upsert');
}

/** Esvazia a base para uma importação nova, sem recolocar o seed. */
function clearAllData() {
  try {
    var data = {
      version: 1,
      cursos: [],
      skipSeedMerge: true
    };
    writeData_(data);
    return { ok: true, total: 0, data: data };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function deleteCurso(codigo, anoVisita, unidade) {
  try {
    var data = readData_();
    data.cursos = data.cursos.filter(function (c) {
      return !(
        String(c.codigo) === String(codigo) &&
        Number(c.anoVisita) === Number(anoVisita) &&
        String(c.unidade || '') === String(unidade || '')
      );
    });
    writeData_(data);
    return { ok: true, total: data.cursos.length, data: data };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function normalizeUnidadeCodigo_(raw) {
  var text = String(raw == null ? '' : raw).trim();
  if (!text) return '';
  var key = text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
  if (key === 'bg' || key.indexOf('bangu') >= 0) return 'BG';
  if (key === 'bs' || key.indexOf('bonsucesso') >= 0) return 'BS';
  if (key === 'cg' || key.indexOf('campogrande') >= 0) return 'CG';
  if (key === 'ead' || key.indexOf('distancia') >= 0) return 'EAD';
  return text.toUpperCase();
}

function normalizeAto_(raw) {
  var text = String(raw == null ? '' : raw).trim().replace(/\s+/g, ' ');
  if (!text) return '';
  var key = text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (key.indexOf('autoriz') === 0) return 'Autorização';
  if (key.indexOf('reconhec') === 0) return 'Reconhecimento';
  if (key.indexOf('renov') === 0) return 'Renovação';
  return text;
}

function normalizeCurso_(raw) {
  var indicadores = {};
  INDICADORES_LISTA.forEach(function (key) {
    var v = raw.indicadores && raw.indicadores[key] !== undefined
      ? raw.indicadores[key]
      : (raw[key] !== undefined ? raw[key] : null);
    indicadores[key] = toNumberOrNull_(v);
  });

  return {
    codigo: raw.codigo != null ? (isFinite(Number(raw.codigo)) ? Number(raw.codigo) : String(raw.codigo)) : null,
    curso: String(raw.curso || '').trim(),
    unidade: normalizeUnidadeCodigo_(raw.unidade),
    anoVisita: toNumberOrNull_(raw.anoVisita),
    atoRegulatorio: normalizeAto_(raw.atoRegulatorio || raw.ato),
    conceitoFinalContinuo: toNumberOrNull_(raw.conceitoFinalContinuo),
    conceitoFinalFaixa: toNumberOrNull_(raw.conceitoFinalFaixa),
    indicadores: indicadores
  };
}

/** No upsert, não apaga nota boa com null/#REF! da planilha. */
function mergeCursoPreserveScores_(existing, incoming) {
  var out = {
    codigo: incoming.codigo != null ? incoming.codigo : existing.codigo,
    curso: incoming.curso || existing.curso,
    unidade: incoming.unidade || existing.unidade,
    anoVisita: incoming.anoVisita != null ? incoming.anoVisita : existing.anoVisita,
    atoRegulatorio: incoming.atoRegulatorio || existing.atoRegulatorio,
    conceitoFinalContinuo: incoming.conceitoFinalContinuo != null
      ? incoming.conceitoFinalContinuo
      : existing.conceitoFinalContinuo,
    conceitoFinalFaixa: incoming.conceitoFinalFaixa != null
      ? incoming.conceitoFinalFaixa
      : existing.conceitoFinalFaixa,
    indicadores: Object.assign({}, existing.indicadores || {})
  };
  Object.keys(incoming.indicadores || {}).forEach(function (k) {
    if (incoming.indicadores[k] != null) out.indicadores[k] = incoming.indicadores[k];
  });
  return out;
}

function toNumberOrNull_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'string') {
    var s = v.trim().replace(',', '.');
    if (!s || s.indexOf('#') === 0 || s.toUpperCase() === 'N/A') return null;
    v = s;
  }
  var n = Number(v);
  return isFinite(n) ? n : null;
}

/**
 * Importa planilha Excel (.xlsx) enviada em base64.
 * Converte via Drive (sem rede externa) e lê as abas disponíveis.
 */
function importExcelBase64(base64, fileName, strategy) {
  var tempFile = null;
  var converted = null;
  try {
    if (!base64) return { ok: false, error: 'Arquivo vazio.' };

    var clean = String(base64).replace(/^data:[^;]+;base64,/, '');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(clean),
      MimeType.MICROSOFT_EXCEL,
      fileName || 'import.xlsx'
    );

    var folder = getOrCreateDataFolder_();
    tempFile = folder.createFile(blob);

    var resource = {
      title: (fileName || 'import') + ' (temp)',
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: folder.getId() }]
    };
    converted = Drive.Files.copy(resource, tempFile.getId());

    var ss = SpreadsheetApp.openById(converted.id);
    var parsed = parseSpreadsheetToCursos_(ss);
    var cursos = parsed.cursos || [];

    if (!cursos.length) {
      return {
        ok: false,
        error: 'Nenhum curso encontrado. Abas lidas: ' +
          (parsed.sheetsFound || []).join(', ') +
          '. Esperado: Painel Conceito e/ou Painel Indicadores (com valores, não #REF!).'
      };
    }

    var result = saveCursos(cursos, strategy || 'upsert');
    result.imported = cursos.length;
    result.diagnostics = parsed.diagnostics;
    result.sheetsFound = parsed.sheetsFound;
    return result;
  } catch (err) {
    return {
      ok: false,
      error: String(err && err.message ? err.message : err) +
        ' — Ative o serviço avançado Drive (Drive API) no Apps Script.'
    };
  } finally {
    try { if (converted && converted.id) Drive.Files.remove(converted.id); } catch (e1) {}
    try { if (tempFile) tempFile.setTrashed(true); } catch (e2) {}
  }
}

function getOrCreateDataFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function parseSpreadsheetToCursos_(ss) {
  var allSheetNames = ss.getSheets().map(function (s) { return s.getName(); });
  var conceitoSheet = findSheet_(ss, [
    'Painel Conceito', 'Conceito', 'painel conceito', 'Painel de Conceito'
  ]);
  var indicadoresSheet = findSheet_(ss, [
    'Painel Indicadores', 'Painel de Indicadores', 'painel indicadores'
  ]);
  // Aba bruta longa (Codigo | Indicador | Nota | ...), comum na origem das fórmulas
  var brutoSheet = findSheet_(ss, ['Indicadores', 'Base Indicadores', 'Dados Indicadores']);

  var byKey = {};
  var diagnostics = {
    sheetsFound: allSheetNames,
    conceitoRows: 0,
    indicadoresRows: 0,
    brutoRows: 0,
    withContinuo: 0,
    withFaixa: 0,
    withIndicadores: 0,
    refErrors: 0,
    emptyScoreRows: 0
  };

  if (conceitoSheet) {
    var cMatrix = readSheetMatrix_(conceitoSheet);
    diagnostics.refErrors += cMatrix.refErrors;
    if (cMatrix.values.length > 1) {
      var cHeaders = normalizeHeaders_(cMatrix.values[0]);
      for (var r = 1; r < cMatrix.values.length; r++) {
        var row = cMatrix.values[r];
        if (isBlank_(row[0])) continue;
        var obj = rowToObject_(cHeaders, row);
        var curso = {
          codigo: firstDefined_(obj.codigo, obj.codigoemec, obj.codigoemeccurso, row[0]),
          curso: firstDefined_(obj.curso, row[1]),
          unidade: firstDefined_(obj.unidade, row[2]),
          anoVisita: firstDefined_(obj.anovista, obj.anovistita, row[3]),
          atoRegulatorio: firstDefined_(obj.atoregulatorio, obj.atoregulatorio, row[4]),
          conceitoFinalContinuo: firstDefined_(
            obj.conceitofinalcontinuo,
            obj.conceitocontinuo,
            row[5]
          ),
          conceitoFinalFaixa: firstDefined_(
            obj.conceitofinalfaixa,
            obj.conceitofaixa,
            obj.cc,
            row[6]
          ),
          indicadores: {}
        };
        var key = courseKey_(curso);
        byKey[key] = normalizeCurso_(curso);
        diagnostics.conceitoRows++;
      }
    }
  }

  if (indicadoresSheet) {
    var iMatrix = readSheetMatrix_(indicadoresSheet);
    diagnostics.refErrors += iMatrix.refErrors;
    if (iMatrix.values.length > 1) {
      var rawHeaders = iMatrix.values[0];
      var mappedHeaders = [];
      for (var hi = 0; hi < rawHeaders.length; hi++) {
        if (hi < 5) {
          mappedHeaders.push(normalizeHeaderKey_(rawHeaders[hi]));
        } else {
          mappedHeaders.push(
            excelHeaderToIndicator_(rawHeaders[hi], iMatrix.rawValues[0][hi])
          );
        }
      }

      for (var ir = 1; ir < iMatrix.values.length; ir++) {
        var irow = iMatrix.values[ir];
        if (isBlank_(irow[0])) continue;
        var meta = {
          codigo: irow[0],
          curso: irow[1],
          unidade: irow[2],
          anoVisita: irow[3],
          atoRegulatorio: irow[4]
        };
        var indicadores = {};
        var hasInd = false;
        for (var ic = 5; ic < mappedHeaders.length; ic++) {
          var ind = mappedHeaders[ic];
          if (!ind || INDICADORES_LISTA.indexOf(String(ind)) < 0) continue;
          var num = toNumberOrNull_(irow[ic]);
          if (num != null) {
            indicadores[ind] = num;
            hasInd = true;
          }
        }

        var key2 = courseKey_(meta);
        if (!byKey[key2]) {
          byKey[key2] = normalizeCurso_({
            codigo: meta.codigo,
            curso: meta.curso,
            unidade: meta.unidade,
            anoVisita: meta.anoVisita,
            atoRegulatorio: meta.atoRegulatorio,
            conceitoFinalContinuo: null,
            conceitoFinalFaixa: null,
            indicadores: indicadores
          });
        } else if (hasInd) {
          byKey[key2].indicadores = Object.assign({}, byKey[key2].indicadores || {}, indicadores);
        }
        diagnostics.indicadoresRows++;
      }
    }
  }

  // Formato longo: Código | Curso | Unidade | Ano | Ato | Indicador | Nota | (Conceito...)
  if (brutoSheet) {
    var bMatrix = readSheetMatrix_(brutoSheet);
    diagnostics.refErrors += bMatrix.refErrors;
    if (bMatrix.values.length > 1) {
      var bHeaders = normalizeHeaders_(bMatrix.values[0]);
      var idxIndicador = findHeaderIndex_(bHeaders, ['indicador', 'indicadores', 'item']);
      var idxNota = findHeaderIndex_(bHeaders, ['nota', 'conceito', 'valor', 'conceitoindicador']);
      var idxCodigo = findHeaderIndex_(bHeaders, ['codigo', 'codigoemec']);
      var idxCurso = findHeaderIndex_(bHeaders, ['curso']);
      var idxUnidade = findHeaderIndex_(bHeaders, ['unidade']);
      var idxAno = findHeaderIndex_(bHeaders, ['anovista', 'ano']);
      var idxAto = findHeaderIndex_(bHeaders, ['atoregulatorio', 'ato']);
      var idxCont = findHeaderIndex_(bHeaders, ['conceitofinalcontinuo', 'conceitocontinuo', 'cccontinuo']);
      var idxFaixa = findHeaderIndex_(bHeaders, ['conceitofinalfaixa', 'conceitofaixa', 'cc']);

      // Heurística posicional se cabeçalhos não baterem
      if (idxCodigo < 0) idxCodigo = 0;
      if (idxIndicador < 0 && bHeaders.length > 5) idxIndicador = 5;
      if (idxNota < 0 && bHeaders.length > 6) idxNota = 6;

      if (idxIndicador >= 0 && idxNota >= 0) {
        for (var br = 1; br < bMatrix.values.length; br++) {
          var brow = bMatrix.values[br];
          if (isBlank_(brow[idxCodigo])) continue;
          var indName = excelHeaderToIndicator_(brow[idxIndicador], bMatrix.rawValues[br][idxIndicador]);
          if (!indName || INDICADORES_LISTA.indexOf(String(indName)) < 0) continue;

          var metaB = {
            codigo: brow[idxCodigo],
            curso: idxCurso >= 0 ? brow[idxCurso] : '',
            unidade: idxUnidade >= 0 ? brow[idxUnidade] : '',
            anoVisita: idxAno >= 0 ? brow[idxAno] : null,
            atoRegulatorio: idxAto >= 0 ? brow[idxAto] : '',
            conceitoFinalContinuo: idxCont >= 0 ? brow[idxCont] : null,
            conceitoFinalFaixa: idxFaixa >= 0 ? brow[idxFaixa] : null,
            indicadores: {}
          };
          metaB.indicadores[indName] = toNumberOrNull_(brow[idxNota]);

          var keyB = courseKey_(metaB);
          if (!byKey[keyB]) {
            byKey[keyB] = normalizeCurso_(metaB);
          } else {
            if (metaB.indicadores[indName] != null) {
              byKey[keyB].indicadores[indName] = metaB.indicadores[indName];
            }
            if ((byKey[keyB].conceitoFinalContinuo == null) && metaB.conceitoFinalContinuo != null) {
              byKey[keyB].conceitoFinalContinuo = toNumberOrNull_(metaB.conceitoFinalContinuo);
            }
            if ((byKey[keyB].conceitoFinalFaixa == null) && metaB.conceitoFinalFaixa != null) {
              byKey[keyB].conceitoFinalFaixa = toNumberOrNull_(metaB.conceitoFinalFaixa);
            }
          }
          diagnostics.brutoRows++;
        }
      }
    }
  }

  var cursos = Object.keys(byKey).map(function (k) { return byKey[k]; });
  cursos.forEach(function (c) {
    if (c.conceitoFinalContinuo != null) diagnostics.withContinuo++;
    if (c.conceitoFinalFaixa != null) diagnostics.withFaixa++;
    var has = Object.keys(c.indicadores || {}).some(function (k) {
      return c.indicadores[k] != null;
    });
    if (has) diagnostics.withIndicadores++;
    else if (c.conceitoFinalFaixa == null && c.conceitoFinalContinuo == null) diagnostics.emptyScoreRows++;
  });

  return {
    cursos: cursos,
    sheetsFound: allSheetNames,
    diagnostics: diagnostics
  };
}

/** Lê display + raw; conta #REF! / erros. */
function readSheetMatrix_(sheet) {
  var range = sheet.getDataRange();
  var display = range.getDisplayValues();
  var raw = range.getValues();
  var refErrors = 0;
  var values = display.map(function (row, ri) {
    return row.map(function (cell, ci) {
      var text = String(cell == null ? '' : cell).trim();
      if (text.indexOf('#') === 0) {
        refErrors++;
        return '';
      }
      // Preferir número cru quando display veio vazio mas raw é número
      var rawCell = raw[ri][ci];
      if ((text === '' || text === '-') && typeof rawCell === 'number' && isFinite(rawCell)) {
        return rawCell;
      }
      return cell;
    });
  });
  return { values: values, rawValues: raw, refErrors: refErrors };
}

function findSheet_(ss, names) {
  var sheets = ss.getSheets();
  for (var n = 0; n < names.length; n++) {
    var want = String(names[n]).toLowerCase().trim();
    for (var s = 0; s < sheets.length; s++) {
      if (String(sheets[s].getName()).toLowerCase().trim() === want) {
        return sheets[s];
      }
    }
  }
  // fallback: contains
  for (var n2 = 0; n2 < names.length; n2++) {
    var token = String(names[n2]).toLowerCase().trim();
    for (var s2 = 0; s2 < sheets.length; s2++) {
      if (String(sheets[s2].getName()).toLowerCase().indexOf(token) >= 0) {
        return sheets[s2];
      }
    }
  }
  return null;
}

function cursosSignature_(cursos) {
  return (cursos || []).map(function (c) {
    return [
      c && c.codigo,
      c && c.curso,
      c && c.unidade,
      c && c.anoVisita,
      c && c.atoRegulatorio
    ].join('|');
  }).join('\n');
}

function courseKey_(c) {
  var codigo = c && c.codigo != null ? String(c.codigo).trim() : '';
  if (codigo && isFinite(Number(codigo))) codigo = String(Number(codigo));
  var ano = c && c.anoVisita != null && c.anoVisita !== ''
    ? String(Number(c.anoVisita))
    : '';
  var unidade = String((c && c.unidade) || '').trim().toUpperCase();
  return [codigo, ano, unidade].join('|');
}

function countNotas_(c) {
  var n = 0;
  var inds = (c && c.indicadores) || {};
  Object.keys(inds).forEach(function (k) {
    if (inds[k] != null && inds[k] !== '') n++;
  });
  return n;
}

/** Descarta cursos com menos de 10 notas preenchidas (cargas parciais e duplicatas incompletas). */
function dropSparseCursos_(cursos) {
  return (cursos || []).filter(function (c) { return countNotas_(c) >= 10; });
}

/** Remove duplicatas pela chave codigo|ano|unidade, mesclando notas. */
function dedupeCursos_(cursos) {
  var map = {};
  var order = [];
  (cursos || []).forEach(function (raw) {
    var c = normalizeCurso_(raw);
    var key = courseKey_(c);
    if (!key || key === '||') return;
    if (!map[key]) {
      map[key] = c;
      order.push(key);
    } else {
      map[key] = mergeCursoPreserveScores_(map[key], c);
    }
  });
  return order.map(function (k) { return map[k]; });
}

function normalizeHeaders_(headers) {
  return headers.map(normalizeHeaderKey_);
}

function normalizeHeaderKey_(h) {
  return String(h || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

function findHeaderIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    for (var c = 0; c < candidates.length; c++) {
      if (headers[i] === candidates[c]) return i;
    }
  }
  return -1;
}

function firstDefined_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v !== null && v !== undefined && String(v).trim() !== '') return v;
  }
  return null;
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

function rowToObject_(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;
    obj[headers[i]] = row[i];
  }
  if (row[0] !== undefined) obj.codigo = obj.codigo || row[0];
  if (row[1] !== undefined) obj.curso = obj.curso || row[1];
  if (row[2] !== undefined) obj.unidade = obj.unidade || row[2];
  if (row[3] !== undefined) obj.anovista = obj.anovista || row[3];
  if (row[4] !== undefined) obj.atoregulatorio = obj.atoregulatorio || row[4];
  if (row[5] !== undefined) obj.conceitofinalcontinuo = obj.conceitofinalcontinuo || row[5];
  if (row[6] !== undefined) obj.conceitofinalfaixa = obj.conceitofinalfaixa || row[6];
  return obj;
}

/**
 * Converte cabeçalho Excel (texto ou data serial mal interpretada) em "1.1", "2.3"...
 */
function excelHeaderToIndicator_(display, raw) {
  if (Object.prototype.toString.call(raw) === '[object Date]' || raw instanceof Date) {
    return raw.getDate() + '.' + (raw.getMonth() + 1);
  }
  var s = String(display || raw || '').trim();
  if (/^\d+\.\d+$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})(?:\/\d{2,4})?$/);
  if (m) return Number(m[1]) + '.' + Number(m[2]);
  // serial date number displayed oddly
  if (typeof raw === 'number' && raw > 1 && raw < 100000) {
    try {
      var d = new Date(Math.round((raw - 25569) * 86400 * 1000));
      if (!isNaN(d.getTime())) return d.getUTCDate() + '.' + (d.getUTCMonth() + 1);
    } catch (e) {}
  }
  return s;
}

/** Utilitário: reinicializa o JSON a partir do seed embutido. */
function resetToSeed() {
  var seed = typeof SEED_DATA !== 'undefined' ? JSON.parse(JSON.stringify(SEED_DATA)) : { version: 1, cursos: [] };
  writeData_(seed);
  return { ok: true, total: seed.cursos.length };
}
