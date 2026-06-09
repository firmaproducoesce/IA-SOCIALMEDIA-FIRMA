const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────
// SYSTEM PROMPT — Especialista em Social Media
// ─────────────────────────────────────────────
const SYSTEM_PROMPT = `Você é um especialista sênior em planejamento de cobertura de eventos para redes sociais, trabalhando para a Firma Produções — uma produtora de eventos de alto nível.

Seu trabalho é criar planejamentos completos, criativos, estratégicos e altamente detalhados para coberturas de eventos nas redes sociais.

RETORNE APENAS UM JSON VÁLIDO e puro, sem markdown, sem blocos de código, sem explicações. Apenas o JSON.

A estrutura deve ser exatamente esta:
{
  "resumo": {
    "evento": "nome do evento",
    "tipo": "tipo do evento",
    "local": "local",
    "objetivo": "objetivo principal da cobertura nas redes",
    "tom_comunicacao": "tom e voz para comunicação",
    "publico_alvo": "descrição detalhada do público-alvo",
    "conceito_criativo": "conceito criativo central que vai guiar toda a cobertura"
  },
  "cronograma": [
    {
      "data": "DD/MM/AAAA",
      "tipo": "montagem ou evento ou pós-evento",
      "periodo": "manhã ou tarde ou noite ou dia todo",
      "atividades_producao": "o que acontece na produção neste período",
      "conteudos_sugeridos": ["conteúdo 1", "conteúdo 2", "conteúdo 3"],
      "plataformas_foco": ["Instagram", "TikTok"],
      "prioridade": "alta ou média ou baixa",
      "observacoes": "observações importantes para este período"
    }
  ],
  "personagens": [
    {
      "nome": "nome completo",
      "papel": "produtor ou cliente ou artista ou parceiro ou outro",
      "como_apresentar": "como introduzir esta pessoa de forma envolvente nas redes sociais",
      "momentos_chave": ["momento ideal 1 para capturar", "momento ideal 2", "momento ideal 3"],
      "pontos_destaque": "o que torna esta pessoa interessante e relevante para o conteúdo"
    }
  ],
  "estrategia_plataformas": [
    {
      "plataforma": "nome da plataforma",
      "estrategia_geral": "estratégia detalhada e específica para esta plataforma",
      "formatos_principais": ["Reels", "Stories", "Feed"],
      "frequencia_posts": "X posts por dia",
      "melhor_horario": "horário ideal para postar (ex: 19h-21h)",
      "hashtags_sugeridas": ["#hashtag1", "#hashtag2", "#hashtag3", "#hashtag4", "#hashtag5"],
      "tom_especifico": "como o tom e estilo de comunicação muda para esta plataforma especificamente"
    }
  ],
  "conteudos": [
    {
      "id": 1,
      "titulo": "título criativo e chamativo do conteúdo",
      "data": "DD/MM",
      "plataforma": "plataforma alvo",
      "formato": "Reels ou Stories ou Post ou TikTok ou YouTube Shorts",
      "hook_5s": "EXATAMENTE o que acontece nos primeiros 5 segundos — seja muito específico: o que o espectador vê, ouve, o texto na tela, o movimento de câmera, a fala de abertura. Este hook deve ser irresistível.",
      "roteiro_completo": "roteiro detalhado de todo o conteúdo: início, desenvolvimento, encerramento. Inclua sugestões de cortes, efeitos, transições.",
      "cta": "call to action específico, direto e poderoso",
      "legenda": "legenda completa pronta para postar, com emojis estratégicos e hashtags relevantes",
      "audio_sugerido": "sugestão concreta de música ou tendência de áudio do momento",
      "objetivo": "o que este conteúdo específico deve gerar (alcance, engajamento, conversão, awareness, etc.)",
      "dica_edicao": "dica específica de edição, efeito, filtro ou técnica que vai elevar este conteúdo"
    }
  ],
  "checklist_producao": [
    {
      "categoria": "Equipamentos ou Logística ou Conteúdo ou Pós-produção ou Estratégia Digital",
      "item": "item específico e acionável do checklist",
      "responsavel": "cargo ou pessoa responsável",
      "prazo": "quando deve estar pronto/concluído"
    }
  ],
  "metricas": {
    "kpis_principais": ["KPI 1 com meta numérica específica", "KPI 2 com meta", "KPI 3"],
    "meta_alcance_estimado": "número estimado de pessoas alcançadas durante o evento",
    "meta_engajamento": "taxa de engajamento esperada (ex: 5-8%)",
    "indicadores_sucesso": ["indicador concreto e mensurável 1", "indicador 2", "indicador 3"],
    "ferramentas_medicao": ["Meta Business Suite", "TikTok Analytics", "etc."]
  }
}

REGRAS IMPORTANTES:
- Seja extremamente criativo, específico e prático. Crie conteúdos que realmente funcionam nas redes sociais de 2025.
- Para os hooks de 5 segundos: seja ULTRA específico. Descreva exatamente o visual, a fala, o texto na tela, o ângulo de câmera.
- Para CTAs: seja direto, urgente e instigante. Nada genérico.
- Gere no mínimo 10 ideias de conteúdo diversas entre as plataformas informadas.
- Adapte tudo ao contexto real do evento. Sem respostas genéricas.
- Use tendências reais de 2025 nas sugestões de formato e áudio.`;

// ─────────────────────────────────────────────
// Formatar dados do evento para o Claude
// ─────────────────────────────────────────────
function formatEventData(data) {
  return `Crie um planejamento completo de cobertura de evento para redes sociais:

EVENTO: ${data.nome || 'Não informado'}
TIPO: ${data.tipo || 'Não informado'}
LOCAL: ${data.local || 'Não informado'}
CLIENTE: ${data.cliente || 'Não informado'}

BRIEFING GERAL:
${data.briefing || 'Não informado'}

OBJETIVO PRINCIPAL: ${data.objetivo || 'Não informado'}

DIAS DE MONTAGEM: ${data.diasMontagem || 'Não informado'}
DIAS DO EVENTO: ${data.diasEvento || 'Não informado'}

PESSOAS IMPORTANTES:
- Produtores: ${data.produtores || 'Não informado'}
- Clientes/Contratantes: ${data.clientes || 'Não informado'}
- Artistas/Performers: ${data.artistas || 'Não informado'}
- Outros: ${data.outros || 'Não informado'}

PLATAFORMAS ALVO: ${(data.plataformas || []).join(', ') || 'Não informado'}

OBSERVAÇÕES ADICIONAIS: ${data.observacoes || 'Nenhuma'}`;
}

// ─────────────────────────────────────────────
// Gerar Excel com ExcelJS
// ─────────────────────────────────────────────
async function generateExcel(planning, eventData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Firma AI';
  workbook.created = new Date();

  // Paleta de cores
  const C = {
    RED: 'FFCC0000',
    RED_BRIGHT: 'FFEE0000',
    RED_LIGHT: 'FFFFECEC',
    BLACK: 'FF111111',
    WHITE: 'FFFFFFFF',
    GRAY_LIGHT: 'FFF7F7F7',
    GRAY_MID: 'FFEEEEEE',
    GRAY_DARK: 'FF555555',
    BORDER: 'FFDDDDDD',
  };

  const titleStyle = (merge) => ({
    font: { bold: true, color: { argb: C.WHITE }, size: 13, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: C.BLACK } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });

  const headerStyle = {
    font: { bold: true, color: { argb: C.WHITE }, size: 10, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: C.RED } },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    border: {
      top: { style: 'thin', color: { argb: C.BLACK } },
      left: { style: 'thin', color: { argb: C.BLACK } },
      bottom: { style: 'thin', color: { argb: C.BLACK } },
      right: { style: 'thin', color: { argb: C.BLACK } },
    },
  };

  const dataStyle = (alt = false) => ({
    font: { color: { argb: C.BLACK }, size: 10, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: alt ? C.GRAY_LIGHT : C.WHITE } },
    alignment: { vertical: 'middle', wrapText: true },
    border: {
      top: { style: 'hair', color: { argb: C.BORDER } },
      left: { style: 'hair', color: { argb: C.BORDER } },
      bottom: { style: 'hair', color: { argb: C.BORDER } },
      right: { style: 'hair', color: { argb: C.BORDER } },
    },
  });

  const applyStyle = (cell, style) => Object.assign(cell, style);

  const addTitleRow = (sheet, text, colCount) => {
    const range = `A1:${String.fromCharCode(64 + colCount)}1`;
    sheet.mergeCells(range);
    const cell = sheet.getCell('A1');
    cell.value = text;
    applyStyle(cell, titleStyle());
    sheet.getRow(1).height = 38;
  };

  // ── Sheet 1: RESUMO ──────────────────────────────────────
  const s1 = workbook.addWorksheet('📋 Resumo', { views: [{ state: 'frozen', ySplit: 2 }] });
  s1.columns = [{ key: 'a', width: 28 }, { key: 'b', width: 75 }];
  addTitleRow(s1, '🎬  FIRMA AI — PLANEJAMENTO DE COBERTURA DE EVENTO', 2);

  const s1Header = s1.addRow(['CAMPO', 'INFORMAÇÃO']);
  s1Header.eachCell(c => applyStyle(c, headerStyle));
  s1.getRow(2).height = 26;

  const resumoRows = [
    ['Nome do Evento', planning.resumo?.evento],
    ['Tipo', planning.resumo?.tipo],
    ['Local', planning.resumo?.local],
    ['Cliente', eventData.cliente],
    ['Objetivo da Cobertura', planning.resumo?.objetivo],
    ['Tom de Comunicação', planning.resumo?.tom_comunicacao],
    ['Público-Alvo', planning.resumo?.publico_alvo],
    ['Conceito Criativo', planning.resumo?.conceito_criativo],
    ['Plataformas', (eventData.plataformas || []).join(', ')],
    ['Gerado em', new Date().toLocaleString('pt-BR')],
  ];

  resumoRows.forEach((item, i) => {
    const row = s1.addRow(item);
    row.height = 36;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    row.getCell(1).font = { bold: true, color: { argb: C.BLACK }, size: 10 };
  });

  // ── Sheet 2: CRONOGRAMA ──────────────────────────────────
  const s2 = workbook.addWorksheet('📅 Cronograma', { views: [{ state: 'frozen', ySplit: 2 }] });
  s2.columns = [
    { key: 'data', width: 14 }, { key: 'tipo', width: 14 }, { key: 'periodo', width: 12 },
    { key: 'atividades', width: 38 }, { key: 'conteudos', width: 45 },
    { key: 'plataformas', width: 22 }, { key: 'prioridade', width: 12 }, { key: 'obs', width: 30 },
  ];
  addTitleRow(s2, '📅  CRONOGRAMA DE PRODUÇÃO E COBERTURA', 8);
  const s2h = s2.addRow(['Data', 'Tipo', 'Período', 'Atividades de Produção', 'Conteúdos Sugeridos', 'Plataformas', 'Prioridade', 'Observações']);
  s2h.eachCell(c => applyStyle(c, headerStyle));
  s2.getRow(2).height = 26;

  (planning.cronograma || []).forEach((item, i) => {
    const conteudos = Array.isArray(item.conteudos_sugeridos)
      ? item.conteudos_sugeridos.map(c => `• ${c}`).join('\n')
      : item.conteudos_sugeridos || '';
    const plats = Array.isArray(item.plataformas_foco) ? item.plataformas_foco.join(', ') : item.plataformas_foco;
    const row = s2.addRow([item.data, item.tipo, item.periodo, item.atividades_producao, conteudos, plats, (item.prioridade || '').toUpperCase(), item.observacoes]);
    row.height = 70;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    const prioCell = row.getCell(7);
    if (item.prioridade === 'alta') prioCell.font = { bold: true, color: { argb: C.RED_BRIGHT }, size: 10 };
    else if (item.prioridade === 'média') prioCell.font = { bold: true, color: { argb: 'FFCC6600' }, size: 10 };
  });

  // ── Sheet 3: PERSONAGENS ─────────────────────────────────
  const s3 = workbook.addWorksheet('👥 Personagens', { views: [{ state: 'frozen', ySplit: 2 }] });
  s3.columns = [
    { key: 'nome', width: 24 }, { key: 'papel', width: 18 },
    { key: 'apresentar', width: 42 }, { key: 'momentos', width: 42 }, { key: 'destaque', width: 38 },
  ];
  addTitleRow(s3, '👥  PERSONAGENS — COMO APRESENTAR NAS REDES', 5);
  const s3h = s3.addRow(['Nome', 'Papel', 'Como Apresentar nas Redes', 'Momentos-Chave para Capturar', 'Pontos de Destaque']);
  s3h.eachCell(c => applyStyle(c, headerStyle));
  s3.getRow(2).height = 26;

  (planning.personagens || []).forEach((p, i) => {
    const momentos = Array.isArray(p.momentos_chave) ? p.momentos_chave.map(m => `• ${m}`).join('\n') : p.momentos_chave;
    const row = s3.addRow([p.nome, p.papel, p.como_apresentar, momentos, p.pontos_destaque]);
    row.height = 70;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    row.getCell(1).font = { bold: true, color: { argb: C.BLACK }, size: 10 };
  });

  // ── Sheet 4: ESTRATÉGIA ──────────────────────────────────
  const s4 = workbook.addWorksheet('📱 Estratégia', { views: [{ state: 'frozen', ySplit: 2 }] });
  s4.columns = [
    { key: 'plat', width: 16 }, { key: 'estrat', width: 48 }, { key: 'formatos', width: 24 },
    { key: 'freq', width: 16 }, { key: 'horario', width: 18 }, { key: 'hashtags', width: 38 }, { key: 'tom', width: 32 },
  ];
  addTitleRow(s4, '📱  ESTRATÉGIA POR PLATAFORMA', 7);
  const s4h = s4.addRow(['Plataforma', 'Estratégia Geral', 'Formatos Principais', 'Frequência', 'Melhor Horário', 'Hashtags Sugeridas', 'Tom Específico']);
  s4h.eachCell(c => applyStyle(c, headerStyle));
  s4.getRow(2).height = 26;

  (planning.estrategia_plataformas || []).forEach((p, i) => {
    const formatos = Array.isArray(p.formatos_principais) ? p.formatos_principais.join(', ') : p.formatos_principais;
    const hashtags = Array.isArray(p.hashtags_sugeridas) ? p.hashtags_sugeridas.join('  ') : p.hashtags_sugeridas;
    const row = s4.addRow([p.plataforma, p.estrategia_geral, formatos, p.frequencia_posts, p.melhor_horario, hashtags, p.tom_especifico]);
    row.height = 90;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    row.getCell(1).font = { bold: true, color: { argb: C.RED }, size: 10 };
  });

  // ── Sheet 5: CONTEÚDOS ───────────────────────────────────
  const s5 = workbook.addWorksheet('🎬 Conteúdos', { views: [{ state: 'frozen', ySplit: 2 }] });
  s5.columns = [
    { key: 'id', width: 5 }, { key: 'titulo', width: 28 }, { key: 'data', width: 10 },
    { key: 'plat', width: 14 }, { key: 'formato', width: 14 },
    { key: 'hook', width: 48 }, { key: 'roteiro', width: 55 },
    { key: 'cta', width: 32 }, { key: 'legenda', width: 42 },
    { key: 'audio', width: 28 }, { key: 'objetivo', width: 28 }, { key: 'edicao', width: 30 },
  ];
  addTitleRow(s5, '🎬  CONTEÚDOS — HOOKS · ROTEIROS · CTAs · LEGENDAS', 12);
  const s5h = s5.addRow(['#', 'Título', 'Data', 'Plataforma', 'Formato', '⚡ HOOK 5 SEGUNDOS', 'Roteiro Completo', 'CTA', 'Legenda', 'Áudio Sugerido', 'Objetivo', 'Dica de Edição']);
  s5h.eachCell((c, col) => {
    applyStyle(c, headerStyle);
    if (col === 6) {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.BLACK } };
    }
  });
  s5.getRow(2).height = 26;

  (planning.conteudos || []).forEach((c, i) => {
    const row = s5.addRow([c.id || i + 1, c.titulo, c.data, c.plataforma, c.formato, c.hook_5s, c.roteiro_completo, c.cta, c.legenda, c.audio_sugerido, c.objetivo, c.dica_edicao]);
    row.height = 110;
    row.eachCell(cell => applyStyle(cell, dataStyle(i % 2 === 0)));
    // Hook cell destacado
    const hookCell = row.getCell(6);
    hookCell.font = { bold: true, color: { argb: C.RED }, size: 10 };
    hookCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.RED_LIGHT } };
    // CTA em negrito
    row.getCell(8).font = { bold: true, color: { argb: C.BLACK }, size: 10 };
  });

  // ── Sheet 6: CHECKLIST ───────────────────────────────────
  const s6 = workbook.addWorksheet('✅ Checklist', { views: [{ state: 'frozen', ySplit: 2 }] });
  s6.columns = [
    { key: 'check', width: 6 }, { key: 'cat', width: 22 },
    { key: 'item', width: 55 }, { key: 'resp', width: 24 }, { key: 'prazo', width: 20 },
  ];
  addTitleRow(s6, '✅  CHECKLIST DE PRODUÇÃO', 5);
  const s6h = s6.addRow(['✓', 'Categoria', 'Item', 'Responsável', 'Prazo']);
  s6h.eachCell(c => applyStyle(c, headerStyle));
  s6.getRow(2).height = 26;

  (planning.checklist_producao || []).forEach((item, i) => {
    const row = s6.addRow(['☐', item.categoria, item.item, item.responsavel, item.prazo]);
    row.height = 28;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(1).font = { size: 13, color: { argb: C.GRAY_DARK } };
  });

  // ── Sheet 7: MÉTRICAS ────────────────────────────────────
  const s7 = workbook.addWorksheet('📊 Métricas', { views: [{ state: 'frozen', ySplit: 2 }] });
  s7.columns = [{ key: 'a', width: 30 }, { key: 'b', width: 65 }];
  addTitleRow(s7, '📊  MÉTRICAS E INDICADORES DE SUCESSO', 2);
  const s7h = s7.addRow(['INDICADOR', 'DETALHES']);
  s7h.eachCell(c => applyStyle(c, headerStyle));
  s7.getRow(2).height = 26;

  const m = planning.metricas || {};
  const metRows = [
    ['Alcance Estimado', m.meta_alcance_estimado],
    ['Meta de Engajamento', m.meta_engajamento],
    ['KPIs Principais', (m.kpis_principais || []).map(k => `• ${k}`).join('\n')],
    ['Indicadores de Sucesso', (m.indicadores_sucesso || []).map(k => `• ${k}`).join('\n')],
    ['Ferramentas de Medição', (m.ferramentas_medicao || []).join(' · ')],
  ];

  metRows.forEach((item, i) => {
    const row = s7.addRow(item);
    row.height = 55;
    row.eachCell(c => applyStyle(c, dataStyle(i % 2 === 0)));
    row.getCell(1).font = { bold: true, color: { argb: C.BLACK }, size: 10 };
  });

  return workbook;
}

// ─────────────────────────────────────────────
// Endpoint principal
// ─────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  try {
    const { apiKey, eventData } = req.body;

    if (!apiKey || !apiKey.startsWith('sk-ant-')) {
      return res.status(400).json({ error: 'API Key inválida. Use uma chave Anthropic válida (começa com sk-ant-).' });
    }

    if (!eventData?.nome || !eventData?.briefing) {
      return res.status(400).json({ error: 'Nome do evento e briefing são obrigatórios.' });
    }

    const client = new Anthropic({ apiKey });
    const userMessage = formatEventData(eventData);

    const response = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 8192,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }],
    });

    let jsonText = response.content[0].text.trim();

    // Limpar possível markdown
    jsonText = jsonText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/i, '').trim();

    // Extrair JSON se houver texto extra
    const jsonMatch = jsonText.match(/\{[\s\S]*\}/);
    if (jsonMatch) jsonText = jsonMatch[0];

    const planning = JSON.parse(jsonText);
    const workbook = await generateExcel(planning, eventData);
    const buffer = await workbook.xlsx.writeBuffer();

    const safeName = (eventData.nome || 'Evento').replace(/[^a-zA-Z0-9À-ÿ\s]/g, '').replace(/\s+/g, '-').trim();
    const filename = `Planejamento-Firma-${safeName}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error('❌ Erro:', err);
    if (err instanceof SyntaxError) {
      res.status(500).json({ error: 'Erro ao processar resposta do Claude. Tente novamente.' });
    } else if (err.status === 401) {
      res.status(401).json({ error: 'API Key inválida ou sem permissão.' });
    } else {
      res.status(500).json({ error: err.message || 'Erro interno do servidor.' });
    }
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════╗
║                                              ║
║      🎬  FIRMA AI EVENTOS — ATIVO            ║
║                                              ║
║   Abra no navegador:                         ║
║   http://localhost:${PORT}                      ║
║                                              ║
╚══════════════════════════════════════════════╝
  `);
});
