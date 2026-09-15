const TOKEN          = 'mrf_rel_9K7xQp2Zv4Ld8Tn3';
const ABA_DESTINO    = 'PEDIDO ATUALIZADO';
const ABA_STAGING    = '_STAGING';
const COLUNAS        = 15;
const COLUNAS_TEXTO  = [];   // ex.: [1, 5] força formato texto (zeros à esquerda)
const MAX_TENTATIVAS = 3;
const ESPERA_MS       = 2000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) return _json({ ok: false, erro: 'token invalido' });

    const props = PropertiesService.getScriptProperties();
    const sessaoAtiva = props.getProperty('sessaoAtiva');

    if (req.bloco === 1) {
      props.setProperty('sessaoAtiva', req.sessionId);
    } else if (sessaoAtiva !== req.sessionId) {
      return _json({ ok: false, erro: 'sessao divergente - outro envio em andamento ou expirado, reinicie do bloco 1' });
    }

    const ss = SpreadsheetApp.getActive();
    let stg = _comRetentativa(function () { return ss.getSheetByName(ABA_STAGING); });

    if (req.bloco === 1) {
      if (stg) _comRetentativa(function () { ss.deleteSheet(stg); });
      stg = _comRetentativa(function () { return ss.insertSheet(ABA_STAGING); });
      COLUNAS_TEXTO.forEach(function (c) {
        stg.getRange(1, c, stg.getMaxRows(), 1).setNumberFormat('@');
      });
      stg.hideSheet();
    }
    if (!stg) return _json({ ok: false, erro: 'staging ausente - reinicie do bloco 1' });

    const dados = req.dados || [];
    if (dados.length) {
      const necessario = req.linhaInicial + dados.length - 1;
      _comRetentativa(function () {
        if (stg.getMaxRows() < necessario) {
          stg.insertRowsAfter(stg.getMaxRows(), necessario - stg.getMaxRows());
        }
        stg.getRange(req.linhaInicial, 1, dados.length, COLUNAS).setValues(dados);
      });
    }

    if (req.bloco < req.totalBlocos) {
      SpreadsheetApp.flush();
      return _json({ ok: true, bloco: req.bloco, linhas: dados.length });
    }

    const total = req.totalLinhas;
    const valores = _comRetentativa(function () {
      return stg.getRange(1, 1, total, COLUNAS).getValues();
    });
    if (valores.length !== total) {
      _comRetentativa(function () { ss.deleteSheet(stg); });
      props.deleteProperty('sessaoAtiva');
      return _json({ ok: false, erro: 'contagem divergente' });
    }

    const dest = ss.getSheetByName(ABA_DESTINO);
    if (!dest) return _json({ ok: false, erro: 'aba destino inexistente' });

    valores[0][0] = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy');

    _comRetentativa(function () {
      if (dest.getMaxRows() < total) dest.insertRowsAfter(dest.getMaxRows(), total - dest.getMaxRows());
      // Limpa só o intervalo realmente usado (nao getMaxRows(), que so cresce
      // ao longo do tempo e encareceria esta operacao a cada envio).
      dest.getRange(1, 1, total, COLUNAS).clearContent();
      if (dest.getMaxRows() > total) {
        dest.deleteRows(total + 1, dest.getMaxRows() - total);
      }
    });

    COLUNAS_TEXTO.forEach(function (c) {
      dest.getRange(1, c, total, 1).setNumberFormat('@');
    });
    dest.getRange(1, 1, 1, 1).setNumberFormat('@');

    _comRetentativa(function () {
      dest.getRange(1, 1, total, COLUNAS).setValues(valores);
    });

    SpreadsheetApp.flush();
    _comRetentativa(function () { ss.deleteSheet(stg); });
    props.deleteProperty('sessaoAtiva');

    return _json({ ok: true, bloco: req.bloco, linhas: total, data: valores[0][0] });
  } catch (err) {
    return _json({ ok: false, erro: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// Reexecuta operacoes do servico Planilhas que podem falhar por timeout
// transitorio de acesso ao documento (lock interno do Google, planilha
// aberta por outro usuario, recalculo de formulas volateis, etc.).
function _comRetentativa(fn) {
  var ultimoErro;
  for (var i = 0; i < MAX_TENTATIVAS; i++) {
    try {
      return fn();
    } catch (err) {
      ultimoErro = err;
      if (i < MAX_TENTATIVAS - 1) Utilities.sleep(ESPERA_MS);
    }
  }
  throw ultimoErro;
}

function _json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
