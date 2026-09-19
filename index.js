'use strict';

const fs = require('fs');
const crypto = require('crypto');
const xlsx = require('xlsx');
const Database = require('better-sqlite3');
const { Resend } = require('resend');

const CONFIG = {
  planilhaOrigem: process.env.PLANILHA_ORIGEM || './notas.xlsx',
  planilhaSaida: process.env.PLANILHA_SAIDA || './notas-processadas.xlsx',
  banco: process.env.BANCO_SQLITE || './notas.db',
  lock: process.env.ARQUIVO_LOCK || './notas.lock',
  abaOrigem: process.env.ABA_ORIGEM || 'Respostas de Estudantes',
  resendApiKey: process.env.RESEND_API_KEY,
  remetente: process.env.EMAIL_REMETENTE,
  destinatario: process.env.EMAIL_DESTINATARIO,
};

const COLUNAS = {
  id: 'activity_id',
  nome: 'nome_da_prova',
  matricula: 'matricula',
  notaEstudante: 'nota_estudante',
  notaMax: 'nota_max_atividade',
  modeloProva: 'modelo_prova_embaralhada',
};

const TIPO = {
  PROCESSADA: 'processada',
  INCONSISTENCIA: 'inconsistencia',
  NAO_PROCESSADO: 'nao_processado',
};

const NAO_IDENTIFICADA = 'Não Identificada';

function normalizar(texto) {
  if (texto === null || texto === undefined) return '';
  return String(texto)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function paraNumero(valor) {
  if (typeof valor === 'number') return Number.isFinite(valor) ? valor : NaN;
  if (valor === null || valor === undefined) return NaN;
  const texto = String(valor).trim();
  if (texto === '') return NaN;
  const normalizado = texto.replace(/\s/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.');
  const numero = Number(normalizado);
  return Number.isFinite(numero) ? numero : NaN;
}

function titleCase(texto) {
  const minusculas = new Set(['de', 'da', 'do', 'das', 'dos', 'e']);
  return normalizar(texto)
    .split(' ')
    .filter(Boolean)
    .map((palavra, i) =>
      i > 0 && minusculas.has(palavra) ? palavra : palavra.charAt(0).toUpperCase() + palavra.slice(1)
    )
    .join(' ');
}

function chaveRegistro(activityId, matricula) {
  return crypto.createHash('sha1').update(`${activityId}|${matricula}`).digest('hex');
}

const REGRAS_SERIE = [
  { re: /\b([123])\s*[ºo°ª]?\s*(?:serie|ano|em)\b/, formato: (m) => `${m[1]}EM` },
  { re: /\b([6-9])\s*[ºo°]?\s*([a-d])\b/, formato: (m) => `${m[1]}${m[2].toUpperCase()}` },
  { re: /\b([6-9])\s*[ºo°]?\s*(?:ano|ef)\b/, formato: (m) => `${m[1]}EF` },
];

function extrairSerie(nomeAtividade) {
  const texto = normalizar(nomeAtividade);
  if (!texto) return NAO_IDENTIFICADA;
  for (const regra of REGRAS_SERIE) {
    const encontrado = texto.match(regra.re);
    if (encontrado) return regra.formato(encontrado);
  }
  return NAO_IDENTIFICADA;
}

const MATERIAS = [
  ['Gramática', /\bgramatica\b/],
  ['Literatura', /\bliteratura\b/],
  ['Português', /\b(?:portugues|lingua portuguesa)\b/],
  ['Filosofia/Sociologia', /\b(?:filosofia|sociologia)\b/],
  ['Matemática', /\b(?:matematica|mat)\b/],
  ['História', /\bhistoria\b/],
  ['Geografia', /\b(?:geografia|geo)\b/],
  ['Física', /\b(?:fisica|fis)\b/],
  ['Química', /\b(?:quimica|quim)\b/],
  ['Biologia', /\b(?:biologia|bio)\b/],
  ['Inglês', /\b(?:ingles|english)\b/],
];

function extrairMateria(nomeAtividade) {
  const texto = normalizar(nomeAtividade);
  if (!texto) return NAO_IDENTIFICADA;
  const principal = texto.split('(')[0];
  for (const [materia, regex] of MATERIAS) {
    if (regex.test(principal)) return materia;
  }
  for (const [materia, regex] of MATERIAS) {
    if (regex.test(texto)) return materia;
  }
  return NAO_IDENTIFICADA;
}

function extrairNumeroLista(nomeAtividade) {
  const texto = normalizar(nomeAtividade);
  if (/\blista\s+extra\b/.test(texto)) return 'Extra';
  const encontrado = texto.match(/\blista\s*(?:n[ºo°]?\s*)?(\d{1,3})\b/);
  return encontrado ? encontrado[1].padStart(2, '0') : null;
}

function extrairFrenteVerso(nomeAtividade) {
  const texto = normalizar(nomeAtividade);
  const frente = texto.match(/\b(?:frente|f)\s*(\d{1,2})\b/);
  if (frente) return `Frente ${frente[1]}`;
  const verso = texto.match(/\bverso\s*(\d{1,2})?\b/);
  if (verso) return verso[1] ? `Verso ${verso[1]}` : 'Verso';
  return null;
}

const PALAVRAS_RESERVADAS = new Set(
  [
    'frente', 'verso', 'recuperacao', 'positivo on', 'lista', 'extra', 'prova',
    'atividade', 'simulado', 'em', 'ef', 'ano', 'serie',
    ...MATERIAS.map(([materia]) => normalizar(materia)),
    'filosofia', 'sociologia', 'portugues', 'gramatica',
  ].map(normalizar)
);

function extrairNomeProfessor(nomeAtividade) {
  const partes = String(nomeAtividade).split(/[-–—]/);
  if (partes.length < 2) return null;
  for (let i = partes.length - 1; i >= 1; i--) {
    const candidato = normalizar(partes[i].replace(/[()]/g, ''));
    if (!candidato || candidato.length < 3) continue;
    if (/\d/.test(candidato)) continue;
    if ([...PALAVRAS_RESERVADAS].some((chave) => candidato.includes(chave))) continue;
    return titleCase(candidato);
  }
  return null;
}

function gerarNomePadronizado(materia, nomeAtividade) {
  if (materia === NAO_IDENTIFICADA) return 'Nome Não Padronizado';
  const numero = extrairNumeroLista(nomeAtividade);
  const frenteVerso = extrairFrenteVerso(nomeAtividade);
  const professor = extrairNomeProfessor(nomeAtividade);

  const partes = [numero ? `Lista ${numero}` : 'Lista', `de ${materia}`];
  if (frenteVerso) partes.push(frenteVerso);

  const base = partes.join(' ');
  return professor ? `${base} - ${professor}` : base;
}

function validarNota(notaBruta, notaMaxBruta, matricula) {
  if (matricula === null || matricula === undefined || String(matricula).trim() === '') {
    return { inconsistente: true, motivo: 'matrícula ausente' };
  }

  const notaMax = paraNumero(notaMaxBruta);
  if (Number.isNaN(notaMax) || notaMax <= 0) {
    return { inconsistente: true, motivo: `nota máxima inválida (${notaMaxBruta})` };
  }

  const vazia = notaBruta === null || notaBruta === undefined || String(notaBruta).trim() === '';
  if (vazia) {
    return { inconsistente: false, realizada: false, nota: null, observacao: '' };
  }

  const nota = paraNumero(notaBruta);
  if (Number.isNaN(nota)) {
    return { inconsistente: true, motivo: `nota não numérica (${notaBruta})` };
  }
  if (nota < 0) {
    return { inconsistente: true, motivo: `nota negativa (${nota})` };
  }
  if (nota <= notaMax) {
    return { inconsistente: false, realizada: true, nota, observacao: '' };
  }

  for (const divisor of [10, 100]) {
    const ajustada = nota / divisor;
    if (ajustada <= notaMax && ajustada > 0) {
      return {
        inconsistente: false,
        realizada: true,
        nota: Number(ajustada.toFixed(4)),
        observacao: `nota original (${nota}) ajustada para ${ajustada}`,
      };
    }
  }

  return { inconsistente: true, motivo: `nota (${nota}) maior que a máxima (${notaMax})` };
}

function abrirBanco() {
  const db = new Database(CONFIG.banco);
  db.pragma('journal_mode = WAL');
  db.pragma('foreign_keys = ON');
  db.exec(`
    CREATE TABLE IF NOT EXISTS registros (
      chave TEXT PRIMARY KEY,
      tipo TEXT NOT NULL,
      activity_id TEXT,
      matricula TEXT,
      nome_atividade TEXT,
      nome_padronizado TEXT,
      modelo_prova TEXT,
      serie TEXT,
      materia TEXT,
      nota REAL,
      nota_max REAL,
      observacao TEXT,
      realizada INTEGER,
      motivo TEXT,
      linha_original TEXT,
      registrado_em TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_registros_tipo ON registros (tipo);

    CREATE TABLE IF NOT EXISTS pendencias (
      chave TEXT PRIMARY KEY,
      matricula TEXT NOT NULL,
      nome_padronizado TEXT NOT NULL,
      detectada_em TEXT NOT NULL,
      notificada_em TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_pendencias_notificada ON pendencias (notificada_em);
  `);
  return db;
}

function lerLinhas() {
  if (!fs.existsSync(CONFIG.planilhaOrigem)) {
    throw new Error(`planilha não encontrada: ${CONFIG.planilhaOrigem}`);
  }

  const workbook = xlsx.readFile(CONFIG.planilhaOrigem, { cellDates: true });
  const aba = workbook.Sheets[CONFIG.abaOrigem];
  if (!aba) {
    throw new Error(`aba '${CONFIG.abaOrigem}' inexistente. Abas: ${workbook.SheetNames.join(', ')}`);
  }

  const linhas = xlsx.utils.sheet_to_json(aba, { defval: '', raw: true });
  if (linhas.length === 0) return { linhas: [], colunasFaltantes: [] };

  const cabecalho = Object.keys(linhas[0]);
  const colunasFaltantes = Object.values(COLUNAS).filter((coluna) => !cabecalho.includes(coluna));
  return { linhas, colunasFaltantes };
}

function classificarLinha(linha) {
  const activityId = String(linha[COLUNAS.id] ?? '').trim();
  const nomeAtividade = String(linha[COLUNAS.nome] ?? '').trim();
  const matricula = String(linha[COLUNAS.matricula] ?? '').trim();
  const base = {
    activity_id: activityId,
    matricula,
    nome_atividade: nomeAtividade,
    linha_original: JSON.stringify(linha),
  };

  if (!activityId || !nomeAtividade) {
    return { ...base, tipo: TIPO.INCONSISTENCIA, motivo: 'activity_id ou nome_da_prova ausente' };
  }

  if (!normalizar(nomeAtividade).includes('lista')) {
    return { ...base, tipo: TIPO.NAO_PROCESSADO, motivo: 'atividade não é uma lista' };
  }

  const resultado = validarNota(linha[COLUNAS.notaEstudante], linha[COLUNAS.notaMax], matricula);
  const serie = extrairSerie(nomeAtividade);
  const materia = extrairMateria(nomeAtividade);

  const motivos = [];
  if (resultado.inconsistente) motivos.push(resultado.motivo);
  if (serie === NAO_IDENTIFICADA) motivos.push('série não identificada');
  if (materia === NAO_IDENTIFICADA) motivos.push('matéria não identificada');

  if (motivos.length > 0) {
    return { ...base, tipo: TIPO.INCONSISTENCIA, motivo: motivos.join('; ') };
  }

  return {
    ...base,
    tipo: TIPO.PROCESSADA,
    nome_padronizado: gerarNomePadronizado(materia, nomeAtividade),
    modelo_prova: String(linha[COLUNAS.modeloProva] ?? ''),
    serie,
    materia,
    nota: resultado.nota,
    nota_max: paraNumero(linha[COLUNAS.notaMax]),
    observacao: resultado.observacao,
    realizada: resultado.realizada ? 1 : 0,
  };
}

function persistir(db, registros) {
  const agora = new Date().toISOString();

  const inserirRegistro = db.prepare(`
    INSERT INTO registros (
      chave, tipo, activity_id, matricula, nome_atividade, nome_padronizado,
      modelo_prova, serie, materia, nota, nota_max, observacao, realizada,
      motivo, linha_original, registrado_em
    ) VALUES (
      @chave, @tipo, @activity_id, @matricula, @nome_atividade, @nome_padronizado,
      @modelo_prova, @serie, @materia, @nota, @nota_max, @observacao, @realizada,
      @motivo, @linha_original, @registrado_em
    )
    ON CONFLICT(chave) DO NOTHING
  `);

  const inserirPendencia = db.prepare(`
    INSERT INTO pendencias (chave, matricula, nome_padronizado, detectada_em, notificada_em)
    VALUES (?, ?, ?, ?, NULL)
    ON CONFLICT(chave) DO NOTHING
  `);

  const transacao = db.transaction((itens) => {
    const resumo = { novos: 0, duplicados: 0, pendencias: 0 };

    for (const item of itens) {
      const chave = chaveRegistro(item.activity_id || item.nome_atividade, item.matricula);
      const resultado = inserirRegistro.run({
        chave,
        tipo: item.tipo,
        activity_id: item.activity_id,
        matricula: item.matricula,
        nome_atividade: item.nome_atividade,
        nome_padronizado: item.nome_padronizado ?? null,
        modelo_prova: item.modelo_prova ?? null,
        serie: item.serie ?? null,
        materia: item.materia ?? null,
        nota: item.nota ?? null,
        nota_max: Number.isNaN(item.nota_max) ? null : item.nota_max ?? null,
        observacao: item.observacao ?? null,
        realizada: item.realizada ?? null,
        motivo: item.motivo ?? null,
        linha_original: item.linha_original,
        registrado_em: agora,
      });

      if (resultado.changes === 0) {
        resumo.duplicados += 1;
        continue;
      }
      resumo.novos += 1;

      if (item.tipo === TIPO.PROCESSADA && item.realizada === 0) {
        const pendencia = inserirPendencia.run(chave, item.matricula, item.nome_padronizado, agora);
        if (pendencia.changes > 0) resumo.pendencias += 1;
      }
    }

    return resumo;
  });

  return transacao(registros);
}

function exportarPlanilha(db) {
  const processadas = db
    .prepare(
      `SELECT activity_id, nome_padronizado, modelo_prova, serie, materia, matricula,
              nota, nota_max, observacao, realizada
         FROM registros WHERE tipo = ? ORDER BY serie, materia, matricula`
    )
    .all(TIPO.PROCESSADA);

  const inconsistencias = db
    .prepare(
      `SELECT activity_id, matricula, nome_atividade, motivo, registrado_em
         FROM registros WHERE tipo = ? ORDER BY registrado_em DESC`
    )
    .all(TIPO.INCONSISTENCIA);

  const naoProcessados = db
    .prepare(
      `SELECT activity_id, matricula, nome_atividade, motivo, registrado_em
         FROM registros WHERE tipo = ? ORDER BY registrado_em DESC`
    )
    .all(TIPO.NAO_PROCESSADO);

  const workbook = xlsx.utils.book_new();

  const linhasProcessadas = processadas.map((r) => ({
    activity_id: r.activity_id,
    nome_padronizado: r.nome_padronizado,
    modelo_prova_embaralhada: r.modelo_prova,
    serie_formatada: r.serie,
    materia_formatada: r.materia,
    matricula: r.matricula,
    nota_estudante: r.realizada ? r.nota : 'Não Realizado',
    nota_max_atividade: r.nota_max,
    observacao: r.observacao,
    realizada: r.realizada ? 'Sim' : 'Não',
  }));

  xlsx.utils.book_append_sheet(
    workbook,
    xlsx.utils.json_to_sheet(linhasProcessadas),
    'PROCESSADAS'
  );
  xlsx.utils.book_append_sheet(
    workbook,
    xlsx.utils.json_to_sheet(inconsistencias),
    'INCONSISTENCIAS'
  );
  xlsx.utils.book_append_sheet(
    workbook,
    xlsx.utils.json_to_sheet(naoProcessados),
    'NAO PROCESSADOS'
  );

  xlsx.writeFile(workbook, CONFIG.planilhaSaida);
  return { processadas: processadas.length, inconsistencias: inconsistencias.length, naoProcessados: naoProcessados.length };
}

function pendenciasPendentesDeAviso(db) {
  return db
    .prepare(
      `SELECT chave, matricula, nome_padronizado
         FROM pendencias WHERE notificada_em IS NULL ORDER BY matricula`
    )
    .all();
}

function marcarNotificadas(db, pendencias) {
  const agora = new Date().toISOString();
  const atualizar = db.prepare('UPDATE pendencias SET notificada_em = ? WHERE chave = ?');
  const transacao = db.transaction((itens) => {
    for (const item of itens) atualizar.run(agora, item.chave);
  });
  transacao(pendencias);
}

function montarHtml(pendencias) {
  const porMatricula = new Map();
  for (const item of pendencias) {
    if (!porMatricula.has(item.matricula)) porMatricula.set(item.matricula, []);
    porMatricula.get(item.matricula).push(item.nome_padronizado);
  }

  const escapar = (texto) =>
    String(texto).replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
    );

  const blocos = [...porMatricula.entries()]
    .map(([matricula, listas]) => {
      const itens = listas.map((l) => `<li>${escapar(l)}</li>`).join('');
      return `<div><h3>Matrícula: ${escapar(matricula)}</h3><ul>${itens}</ul></div>`;
    })
    .join('');

  return `<h2>Alerta de atividades não realizadas</h2><p>${porMatricula.size} matrícula(s), ${pendencias.length} pendência(s).</p><hr>${blocos}`;
}

async function enviarRelatorio(pendencias) {
  if (!CONFIG.resendApiKey || !CONFIG.remetente || !CONFIG.destinatario) {
    throw new Error('RESEND_API_KEY, EMAIL_REMETENTE e EMAIL_DESTINATARIO precisam estar definidos');
  }

  const resend = new Resend(CONFIG.resendApiKey);
  const { error } = await resend.emails.send({
    from: CONFIG.remetente,
    to: [CONFIG.destinatario],
    subject: `Alerta de pendências — ${pendencias.length} item(ns)`,
    html: montarHtml(pendencias),
  });

  if (error) throw new Error(`falha no envio via Resend: ${error.message}`);
}

function adquirirLock() {
  try {
    const descritor = fs.openSync(CONFIG.lock, 'wx');
    fs.writeSync(descritor, String(process.pid));
    return descritor;
  } catch (erro) {
    if (erro.code === 'EEXIST') {
      throw new Error(`execução já em andamento (${CONFIG.lock}). Remova o arquivo se for resíduo.`);
    }
    throw erro;
  }
}

function liberarLock(descritor) {
  try {
    fs.closeSync(descritor);
    fs.unlinkSync(CONFIG.lock);
  } catch (erro) {
    console.error('não foi possível liberar o lock:', erro.message);
  }
}

async function main() {
  const descritorLock = adquirirLock();
  let db;

  try {
    const { linhas, colunasFaltantes } = lerLinhas();
    if (colunasFaltantes.length > 0) {
      throw new Error(`colunas ausentes no cabeçalho: ${colunasFaltantes.join(', ')}`);
    }
    if (linhas.length === 0) {
      console.log('planilha sem linhas de dados.');
      return;
    }

    db = abrirBanco();

    const registros = linhas.map((linha) => {
      try {
        return classificarLinha(linha);
      } catch (erro) {
        return {
          tipo: TIPO.INCONSISTENCIA,
          activity_id: String(linha[COLUNAS.id] ?? ''),
          matricula: String(linha[COLUNAS.matricula] ?? ''),
          nome_atividade: String(linha[COLUNAS.nome] ?? ''),
          motivo: `erro inesperado: ${erro.message}`,
          linha_original: JSON.stringify(linha),
        };
      }
    });

    const resumo = persistir(db, registros);
    const exportado = exportarPlanilha(db);

    const pendencias = pendenciasPendentesDeAviso(db);
    if (pendencias.length > 0) {
      await enviarRelatorio(pendencias);
      marcarNotificadas(db, pendencias);
    }

    console.log(
      [
        `linhas lidas: ${linhas.length}`,
        `novos: ${resumo.novos}`,
        `já conhecidos: ${resumo.duplicados}`,
        `pendências novas: ${resumo.pendencias}`,
        `e-mails enviados: ${pendencias.length > 0 ? 1 : 0}`,
        `planilha gerada: ${CONFIG.planilhaSaida} (${exportado.processadas} processadas, ${exportado.inconsistencias} inconsistências, ${exportado.naoProcessados} não processados)`,
      ].join(' | ')
    );
  } finally {
    if (db) db.close();
    liberarLock(descritorLock);
  }
}

main().catch((erro) => {
  console.error('falha na execução:', erro.message);
  process.exitCode = 1;
});