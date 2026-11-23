const pptxgen = require('/tmp/ambev_presentation/node_modules/pptxgenjs');
const expandScenarios = process.env.AMBEV_EXPAND_SCENARIOS === '1';
const outputFile = process.env.AMBEV_OUTPUT_FILE || 'APRESENTACAO_AMBEV_EXECUTIVA.pptx';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Grupo LABDATA';
pptx.subject = 'Cenários econômicos 2027 — Ambev';
pptx.title = 'Ambev em 2027 — cenários, impacto e estratégia';
pptx.company = 'LABDATA';
pptx.lang = 'pt-BR';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'pt-BR'
};
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.margin = 0;

const C = {
  navy: '10273D',
  ink: '153247',
  blue: '287FA8',
  teal: '0D8C88',
  green: '3A9D5D',
  amber: 'D99B2B',
  orange: 'E87525',
  red: 'C64D49',
  cream: 'F7F5F0',
  paper: 'FFFFFF',
  mist: 'EAF0F3',
  gray: '657582',
  lightGray: 'D9E1E6',
  white: 'FFFFFF'
};

function shape(slide, type, x, y, w, h, fill, line = fill, radius = 0) {
  const shp = radius ? pptx.ShapeType.roundRect : type;
  slide.addShape(shp, { x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { color: line, transparency: line === fill ? 100 : 0 } });
}

function text(slide, value, x, y, w, h, opt = {}) {
  slide.addText(value, {
    x, y, w, h,
    fontFace: opt.fontFace || 'Aptos',
    fontSize: opt.fontSize || 14,
    color: opt.color || C.ink,
    bold: opt.bold || false,
    breakLine: false,
    margin: opt.margin === undefined ? 0 : opt.margin,
    valign: opt.valign || 'mid',
    align: opt.align || 'left',
    fit: 'shrink',
    paraSpaceAfterPt: 0,
    bullet: opt.bullet,
    italic: opt.italic || false,
    transparency: opt.transparency || 0,
    isTextBox: true
  });
}

function line(slide, x1, y1, x2, y2, color, width = 1) {
  slide.addShape(pptx.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width } });
}

function header(slide, kicker, title, subtitle = '') {
  text(slide, kicker.toUpperCase(), 0.58, 0.35, 4.4, 0.2, { fontSize: 8.5, color: C.teal, bold: true, charSpacing: 1.5 });
  text(slide, title, 0.58, 0.60, 11.85, 0.45, { fontSize: 23, color: C.navy, bold: true });
  if (subtitle) text(slide, subtitle, 0.58, 1.08, 11.9, 0.28, { fontSize: 10.5, color: C.gray });
  line(slide, 0.58, 1.43, 12.75, 1.43, C.lightGray, 0.9);
}

function footer(slide, value, page, dark = false) {
  const clr = dark ? 'D7E2E8' : '71818C';
  line(slide, 0.58, 7.07, 12.75, 7.07, dark ? '345166' : C.lightGray, 0.6);
  text(slide, value, 0.58, 7.13, 11.45, 0.17, { fontSize: 6.7, color: clr });
  text(slide, String(page).padStart(2, '0'), 12.25, 7.10, 0.5, 0.18, { fontSize: 7, color: clr, bold: true, align: 'right' });
}

function metricCard(slide, x, y, w, label, value, note, accent) {
  shape(slide, pptx.ShapeType.rect, x, y, w, 1.16, C.white, C.white, 0.10);
  shape(slide, pptx.ShapeType.rect, x, y, 0.07, 1.16, accent, accent);
  text(slide, label.toUpperCase(), x + 0.22, y + 0.14, w - 0.34, 0.16, { fontSize: 7.3, color: C.gray, bold: true, charSpacing: 0.8 });
  text(slide, value, x + 0.22, y + 0.37, w - 0.34, 0.32, { fontSize: 21, color: C.navy, bold: true });
  text(slide, note, x + 0.22, y + 0.79, w - 0.34, 0.20, { fontSize: 8.2, color: C.gray });
}

function bullet(slide, value, x, y, w, color = C.ink) {
  shape(slide, pptx.ShapeType.ellipse, x, y + 0.075, 0.075, 0.075, C.teal, C.teal);
  text(slide, value, x + 0.18, y, w - 0.18, 0.30, { fontSize: 11, color });
}

function scenarioCard(slide, x, color, titleValue, tag, metrics, economy, ambev, action, signal) {
  shape(slide, pptx.ShapeType.rect, x, 1.72, 3.83, 4.93, C.white, C.lightGray, 0.12);
  shape(slide, pptx.ShapeType.rect, x, 1.72, 3.83, 0.14, color, color);
  text(slide, tag.toUpperCase(), x + 0.26, 1.91, 2.3, 0.16, { fontSize: 7.5, color, bold: true, charSpacing: 0.8 });
  text(slide, titleValue, x + 0.26, 2.16, 3.23, 0.31, { fontSize: 17, color: C.navy, bold: true });
  text(slide, metrics, x + 0.26, 2.54, 3.25, 0.46, { fontSize: 8.7, color: C.gray, valign: 'top' });
  line(slide, x + 0.26, 3.13, x + 3.52, 3.13, C.lightGray, 0.6);
  text(slide, 'ECONOMIA', x + 0.26, 3.33, 1.1, 0.14, { fontSize: 7, color: C.gray, bold: true, charSpacing: 0.8 });
  text(slide, economy, x + 0.26, 3.53, 3.20, 0.44, { fontSize: 9.3, color: C.ink, valign: 'top' });
  text(slide, 'AMBEV', x + 0.26, 4.16, 0.8, 0.14, { fontSize: 7, color: C.gray, bold: true, charSpacing: 0.8 });
  text(slide, ambev, x + 0.26, 4.36, 3.20, 0.44, { fontSize: 9.3, color: C.ink, valign: 'top' });
  text(slide, 'AÇÃO', x + 0.26, 5.03, 0.8, 0.14, { fontSize: 7, color: C.gray, bold: true, charSpacing: 0.8 });
  text(slide, action, x + 0.26, 5.23, 3.20, 0.38, { fontSize: 9.3, color: C.ink, valign: 'top' });
  shape(slide, pptx.ShapeType.roundRect, x + 0.23, 5.85, 3.30, 0.49, C.cream, C.cream, 0.08);
  text(slide, 'GATILHO  ' + signal, x + 0.36, 5.97, 3.05, 0.18, { fontSize: 7.7, color: C.ink, bold: true });
}

// Slide 1
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  shape(s, pptx.ShapeType.rect, 0, 0, 13.333, 7.5, C.navy, C.navy);
  shape(s, pptx.ShapeType.arc, 9.35, -1.10, 5.2, 5.2, '173A54', '173A54');
  shape(s, pptx.ShapeType.arc, 10.60, 3.45, 3.8, 3.8, '1C4961', '1C4961');
  text(s, 'ECONOMIA APLICADA  |  CENÁRIOS 2027', 0.68, 0.57, 5.5, 0.20, { fontSize: 8.5, color: '92C9C4', bold: true, charSpacing: 1.4 });
  text(s, 'Ambev entra em 2027 forte.\nO desafio é defender volume e margem.', 0.68, 1.08, 8.15, 1.16, { fontSize: 28, color: C.white, bold: true, valign: 'top' });
  text(s, 'A companhia combina caixa líquido e hedge cambial, mas ainda depende do consumo doméstico e enfrenta custos, logística e tributação pressionados.', 0.70, 2.47, 6.7, 0.58, { fontSize: 12.5, color: 'D7E2E8', valign: 'top' });
  shape(s, pptx.ShapeType.roundRect, 0.68, 3.45, 6.56, 1.12, '173A54', '173A54', 0.12);
  text(s, 'TESE DO GRUPO', 0.95, 3.69, 1.35, 0.16, { fontSize: 7.5, color: '92C9C4', bold: true, charSpacing: 1.0 });
  text(s, 'No cenário-base, preservar participação e eficiência vem antes de ampliar capex.', 0.95, 3.98, 5.75, 0.28, { fontSize: 14.2, color: C.white, bold: true });
  metricCard(s, 0.68, 5.21, 2.70, 'Receita líquida 2025', 'R$ 88,2 bi', 'Escala financeira', '87C7C2');
  metricCard(s, 3.57, 5.21, 2.70, 'Caixa líquido 2025', 'R$ 16,9 bi', 'Resiliência a juros altos', 'F1B44C');
  metricCard(s, 6.46, 5.21, 2.70, 'Brasil no 3T26', '57,2%', 'Da receita consolidada', 'F08A4B');
  text(s, 'AMBEV  /  ABEV3', 9.55, 5.58, 2.65, 0.28, { fontSize: 14, color: 'D7E2E8', bold: true, align: 'right' });
  text(s, 'Posicionamento para os cenários do enunciado', 9.45, 5.97, 2.75, 0.20, { fontSize: 8.2, color: '92C9C4', align: 'right' });
  footer(s, 'Fontes: Ambev, Release 3T26 (29/10/2026); Form 20-F 2025; RI Ambev — Dados consultados em 23/11/2026.', 1, true);
}

// Slide 2
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Leitura da empresa', 'A tese depende de três variáveis — e não da Selic isoladamente.', 'A exposição da Ambev é operacional: consumo, custo por hectolitro e tributação.');
  const x = [0.58, 4.55, 8.52];
  const color = [C.teal, C.orange, C.amber];
  const labels = ['1  DEMANDA', '2  MARGEM', '3  REGRAS DO JOGO'];
  const titles = ['Renda e volume', 'Preço / mix vs. custo', 'Câmbio e tributação'];
  const body = [
    'PIB, emprego, renda real e crédito definem ocasiões de consumo e elasticidade a preço.',
    'O indicador-chave é receita/hl versus CPV/hl: repasse só cria valor se não destruir volume.',
    'BRL/USD, commodities, diesel e o Imposto Seletivo de 2027 determinam parte relevante da pressão de custo.'
  ];
  for (let i = 0; i < 3; i++) {
    shape(s, pptx.ShapeType.roundRect, x[i], 1.83, 3.63, 2.32, C.white, C.white, 0.12);
    shape(s, pptx.ShapeType.ellipse, x[i] + 0.26, 2.10, 0.48, 0.48, color[i], color[i]);
    text(s, String(i + 1), x[i] + 0.26, 2.17, 0.48, 0.16, { fontSize: 10, color: C.white, bold: true, align: 'center' });
    text(s, labels[i], x[i] + 0.90, 2.03, 2.35, 0.15, { fontSize: 7.2, color: C.gray, bold: true, charSpacing: 0.8 });
    text(s, titles[i], x[i] + 0.90, 2.28, 2.35, 0.25, { fontSize: 15, color: C.navy, bold: true });
    text(s, body[i], x[i] + 0.28, 2.84, 3.05, 0.82, { fontSize: 10.2, color: C.ink, valign: 'top' });
  }
  shape(s, pptx.ShapeType.roundRect, 0.58, 4.56, 12.17, 1.78, C.navy, C.navy, 0.12);
  text(s, 'O QUE O 3T26 JÁ MOSTRA', 0.88, 4.86, 2.65, 0.17, { fontSize: 8, color: '92C9C4', bold: true, charSpacing: 1.1 });
  text(s, 'Cerveja Brasil', 0.88, 5.33, 1.55, 0.19, { fontSize: 10, color: C.white, bold: true });
  text(s, '+0,5%', 0.88, 5.64, 1.30, 0.30, { fontSize: 24, color: '87C7C2', bold: true });
  text(s, 'volume', 2.12, 5.71, 0.68, 0.16, { fontSize: 8, color: 'D7E2E8' });
  line(s, 3.12, 5.12, 3.12, 6.02, '345166', 0.8);
  text(s, 'Receita/hl', 3.47, 5.33, 1.25, 0.19, { fontSize: 10, color: C.white, bold: true });
  text(s, '+4,8%', 3.47, 5.64, 1.32, 0.30, { fontSize: 24, color: 'F1B44C', bold: true });
  line(s, 5.75, 5.12, 5.75, 6.02, '345166', 0.8);
  text(s, 'CPV/hl caixa', 6.10, 5.33, 1.45, 0.19, { fontSize: 10, color: C.white, bold: true });
  text(s, '+5,4%', 6.10, 5.64, 1.32, 0.30, { fontSize: 24, color: 'F08A4B', bold: true });
  line(s, 8.74, 5.12, 8.74, 6.02, '345166', 0.8);
  text(s, 'Proteção 2027', 9.08, 5.33, 1.57, 0.19, { fontSize: 10, color: C.white, bold: true });
  text(s, 'Hedge BRL/USD', 9.08, 5.64, 2.35, 0.24, { fontSize: 17, color: '87C7C2', bold: true });
  text(s, 'cobre exposição cambial; não todo o choque de custo', 9.08, 5.98, 2.82, 0.15, { fontSize: 7.5, color: 'D7E2E8' });
  footer(s, 'Fontes: Ambev, Release 3T26; Form 20-F 2025 (insumos, câmbio e reforma tributária).', 2);
}

// Slide 3 — matriz original ou três slides detalhados de cenários
if (!expandScenarios) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Cenários 2027', 'O cenário-base exige proteção de margem sem sacrificar volume.', 'Cenários do enunciado; impactos e ações são análise do grupo a partir das fontes verificadas.');
  scenarioCard(s, 0.58, C.green, 'Ajuste com crescimento', 'Otimista', 'PIB 2,5%–3,0% | IPCA 3,5%–4,0%\nSelic 10,5%–11,5% | BRL 4,80–5,10', 'Crédito, renda e atividade aceleram; inflação e logística aliviam.', 'Volume, consumo fora do lar e premiumização ganham força; custos recuam.', 'Ampliar capex comercial, premium e canais digitais.', 'Volume sobe e CPV/hl < receita/hl');
  scenarioCard(s, 4.75, C.amber, 'Equilíbrio instável', 'Base', 'PIB 1,2%–1,8% | IPCA 4,3%–5,0%\nSelic 12%–13% | BRL 5,10–5,50', 'Juros altos sustentam crédito fraco e câmbio volátil.', 'Preço/mix protege receita, mas volume e margem ficam pressionados.', 'Capex seletivo; retornáveis; produtividade e participação.', 'Receita/hl cresce sem volume');
  scenarioCard(s, 8.92, C.red, 'Choque duplo', 'Pessimista', 'IPCA 6,0%–7,5% | Selic 15%–16%\nBRL 5,80–6,50 | Brent > US$ 120', 'Inflação, petróleo e risco fiscal reduzem renda e crédito.', 'Insumos, frete e embalagem sobem; demanda fica mais sensível a preço.', 'Caixa, eficiência, fornecedores alternativos e acessibilidade.', 'Real/diesel sobem + volumes caem');
  footer(s, 'Fontes: enunciado do trabalho; Ambev, Release 3T26 e Form 20-F 2025; MF/SPE, Boletim Macrofiscal (23/11/2026).', 3);
}

if (expandScenarios) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Cenários 2027  |  1 de 3', 'Ajuste com crescimento: capturar volume sem trocar margem por preço.', 'Cenário otimista do enunciado — a alavanca central é a recuperação da demanda doméstica.');
  shape(s, pptx.ShapeType.roundRect, 0.58, 1.77, 3.20, 4.91, C.green, C.green, 0.12);
  text(s, 'OTIMISTA', 0.92, 2.10, 1.35, 0.18, { fontSize: 8, color: 'E6F4EA', bold: true, charSpacing: 1.1 });
  text(s, 'Ajuste com\ncrescimento', 0.92, 2.47, 2.35, 0.68, { fontSize: 22, color: C.white, bold: true, valign: 'top' });
  line(s, 0.92, 3.38, 3.18, 3.38, '7AC18D', 0.8);
  text(s, 'PREMISSAS DO ENUNCIADO', 0.92, 3.67, 2.35, 0.15, { fontSize: 7, color: 'E6F4EA', bold: true, charSpacing: 0.7 });
  text(s, 'PIB 2,5%–3,0%\nIPCA 3,5%–4,0%\nSelic 10,5%–11,5%\nBRL 4,80–5,10\nBrent ~ US$ 75', 0.92, 4.01, 1.85, 1.22, { fontSize: 14, color: C.white, bold: true, valign: 'top' });
  text(s, 'Economia: crédito e renda ganham tração; inflação e logística aliviam.', 0.92, 5.72, 2.22, 0.47, { fontSize: 9.1, color: 'E6F4EA', valign: 'top' });
  const blocks = [
    ['IMPLICAÇÃO PARA A AMBEV', 'Maior consumo fora do lar e renda disponível favorecem volume, premiumização e NAB. Câmbio e energia mais benignos aliviam insumos e distribuição.'],
    ['DECISÃO', 'Antecipar capex comercial e de distribuição onde o volume responder; acelerar premium, retornáveis e canais digitais.'],
    ['GATILHO PARA ACELERAR', 'Volume de Cerveja Brasil e NAB cresce com CPV/hl abaixo da receita/hl.']
  ];
  blocks.forEach((b, i) => {
    const y = 1.86 + i * 1.47;
    shape(s, pptx.ShapeType.roundRect, 4.12, y, 8.63, 1.13, C.white, C.white, 0.09);
    shape(s, pptx.ShapeType.rect, 4.12, y, 0.10, 1.13, i === 1 ? C.teal : C.green, i === 1 ? C.teal : C.green);
    text(s, b[0], 4.46, y + 0.20, 2.55, 0.15, { fontSize: 7.4, color: C.gray, bold: true, charSpacing: 0.8 });
    text(s, b[1], 4.46, y + 0.47, 7.72, 0.38, { fontSize: 11.1, color: C.ink, bold: i === 2, valign: 'top' });
  });
  shape(s, pptx.ShapeType.roundRect, 4.12, 6.27, 8.63, 0.41, 'E6F4EA', 'E6F4EA', 0.06);
  text(s, 'LEITURA CRÍTICA  |  O ganho mais relevante é operacional — volume e margem —, não financeiro.', 4.43, 6.39, 7.85, 0.14, { fontSize: 8.2, color: C.green, bold: true });
  footer(s, 'Fontes: enunciado do trabalho; Ambev, Release 3T26 e Form 20-F 2025. Impactos e decisões: análise do grupo.', 3);
}

if (expandScenarios) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Cenários 2027  |  2 de 3', 'Equilíbrio instável: proteger participação e margem em ambiente ainda restritivo.', 'Cenário-base do grupo — mais aderente à projeção macro oficial e aos sinais do 3T26.');
  shape(s, pptx.ShapeType.roundRect, 0.58, 1.77, 3.20, 4.91, C.amber, C.amber, 0.12);
  text(s, 'CENÁRIO-BASE', 0.92, 2.10, 1.75, 0.18, { fontSize: 8, color: 'FFF1D5', bold: true, charSpacing: 1.0 });
  text(s, 'Equilíbrio\ninstável', 0.92, 2.47, 2.30, 0.68, { fontSize: 22, color: C.white, bold: true, valign: 'top' });
  line(s, 0.92, 3.38, 3.18, 3.38, 'E9C171', 0.8);
  text(s, 'REFERÊNCIA PARA 2027', 0.92, 3.67, 2.35, 0.15, { fontSize: 7, color: 'FFF1D5', bold: true, charSpacing: 0.7 });
  text(s, 'SPE/MF:\nPIB 1,8%\nIPCA 4,5%\nBRL médio 5,34\nSelic acumulada 12,93%', 0.92, 4.01, 2.15, 1.30, { fontSize: 13.5, color: C.white, bold: true, valign: 'top' });
  text(s, 'Enunciado: PIB 1,2%–1,8%; Selic 12%–13%; BRL 5,10–5,50.', 0.92, 5.72, 2.22, 0.47, { fontSize: 8.7, color: 'FFF1D5', valign: 'top' });
  const blocks = [
    ['O QUE JÁ É VISÍVEL', 'No 3T26, Cerveja Brasil cresceu 0,5% em volume; receita/hl avançou 4,8%, abaixo do CPV/hl caixa (+5,4%).'],
    ['IMPLICAÇÃO PARA A AMBEV', 'Preço/mix sustenta receita, mas juros, renda mais fraca, câmbio e frete limitam volume e pressionam margem. O hedge BRL/USD reduz apenas parte do risco.'],
    ['DECISÃO E ALERTA', 'Priorizar retornáveis, portfólio acessível, produtividade e capex seletivo. Alerta: receita/hl cresce sem recuperação de volume ou tributação de bebidas se torna mais onerosa.']
  ];
  blocks.forEach((b, i) => {
    const y = 1.86 + i * 1.47;
    shape(s, pptx.ShapeType.roundRect, 4.12, y, 8.63, 1.13, C.white, C.white, 0.09);
    shape(s, pptx.ShapeType.rect, 4.12, y, 0.10, 1.13, i === 2 ? C.orange : C.amber, i === 2 ? C.orange : C.amber);
    text(s, b[0], 4.46, y + 0.20, 2.70, 0.15, { fontSize: 7.4, color: C.gray, bold: true, charSpacing: 0.8 });
    text(s, b[1], 4.46, y + 0.47, 7.72, 0.40, { fontSize: 10.65, color: C.ink, bold: i === 2, valign: 'top' });
  });
  shape(s, pptx.ShapeType.roundRect, 4.12, 6.27, 8.63, 0.41, 'FFF1D5', 'FFF1D5', 0.06);
  text(s, 'RISCO REGULATÓRIO  |  Imposto Seletivo sobre bebidas alcoólicas e açucaradas entra em vigor em 2027; alíquotas ainda não estavam plenamente definidas.', 4.43, 6.35, 7.92, 0.16, { fontSize: 7.7, color: C.amber, bold: true });
  footer(s, 'Fontes: MF/SPE, Boletim Macrofiscal (23/11/2026); Ambev, Release 3T26 e Form 20-F 2025; enunciado.', 4);
}

if (expandScenarios) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Cenários 2027  |  3 de 3', 'Choque duplo: preservar liquidez e acessibilidade sem abandonar participação.', 'Cenário pessimista — o choque combina desvalorização cambial, inflação e contração do crédito.');
  shape(s, pptx.ShapeType.roundRect, 0.58, 1.77, 3.20, 4.91, C.red, C.red, 0.12);
  text(s, 'PESSIMISTA', 0.92, 2.10, 1.40, 0.18, { fontSize: 8, color: 'FCE5E3', bold: true, charSpacing: 1.1 });
  text(s, 'Choque\nduplo', 0.92, 2.47, 2.30, 0.68, { fontSize: 22, color: C.white, bold: true, valign: 'top' });
  line(s, 0.92, 3.38, 3.18, 3.38, 'E49A96', 0.8);
  text(s, 'PREMISSAS DO ENUNCIADO', 0.92, 3.67, 2.35, 0.15, { fontSize: 7, color: 'FCE5E3', bold: true, charSpacing: 0.7 });
  text(s, 'IPCA 6,0%–7,5%\nSelic 15%–16%\nBRL 5,80–6,50\nBrent > US$ 120\nCrédito encolhe', 0.92, 4.01, 2.10, 1.22, { fontSize: 13.5, color: C.white, bold: true, valign: 'top' });
  text(s, 'Economia: renda e crédito retraem; inflação e câmbio elevam o custo de vida.', 0.92, 5.72, 2.22, 0.47, { fontSize: 9.0, color: 'FCE5E3', valign: 'top' });
  const blocks = [
    ['CANAIS DE PRESSÃO', 'Dólar, commodities e petróleo encarecem embalagem, insumos e distribuição; perda de renda aumenta a sensibilidade do consumidor ao preço.'],
    ['RISCO RESIDUAL', 'Hedge BRL/USD protege uma exposição definida, mas não neutraliza custo de commodities, frete, energia ou queda de demanda. A LAS adiciona risco de inflação e logística.'],
    ['DECISÃO E ALERTA', 'Preservar caixa, adiar capex não essencial, reforçar alternativas de fornecedor e faixas de preço acessíveis. Alerta: real/diesel sobem e os dois volumes Brasil caem.']
  ];
  blocks.forEach((b, i) => {
    const y = 1.86 + i * 1.47;
    shape(s, pptx.ShapeType.roundRect, 4.12, y, 8.63, 1.13, C.white, C.white, 0.09);
    shape(s, pptx.ShapeType.rect, 4.12, y, 0.10, 1.13, i === 2 ? C.orange : C.red, i === 2 ? C.orange : C.red);
    text(s, b[0], 4.46, y + 0.20, 2.45, 0.15, { fontSize: 7.4, color: C.gray, bold: true, charSpacing: 0.8 });
    text(s, b[1], 4.46, y + 0.47, 7.72, 0.40, { fontSize: 10.65, color: C.ink, bold: i === 2, valign: 'top' });
  });
  shape(s, pptx.ShapeType.roundRect, 4.12, 6.27, 8.63, 0.41, 'FCE5E3', 'FCE5E3', 0.06);
  text(s, 'PRIORIDADE  |  Evitar que a defesa de margem via preço acelere a perda de volume e participação.', 4.43, 6.39, 7.70, 0.14, { fontSize: 8.15, color: C.red, bold: true });
  footer(s, 'Fontes: enunciado do trabalho; Ambev, Form 20-F 2025 (câmbio, insumos, riscos operacionais); Release 3T26.', 5);
}

// Slide 4
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Recomendação estratégica', 'Núcleo robusto: proteger participação, caixa e flexibilidade.', 'A solidez financeira permite escolher onde acelerar — não elimina os riscos operacionais.');
  shape(s, pptx.ShapeType.roundRect, 0.58, 1.77, 7.15, 4.85, C.white, C.white, 0.12);
  text(s, 'O QUE FAZER EM QUALQUER CENÁRIO', 0.90, 2.08, 3.45, 0.16, { fontSize: 8, color: C.teal, bold: true, charSpacing: 1.0 });
  const actions = [
    ['01', 'Defender volume com arquitetura de preço', 'Retornáveis, portfólio acessível e execução no ponto de venda.'],
    ['02', 'Travar eficiência de custo', 'Contratos de insumo, produtividade fabril e logística.'],
    ['03', 'Preservar opcionalidade financeira', 'Capex disciplinado e caixa para responder a oportunidades.']
  ];
  actions.forEach((a, i) => {
    const y = 2.56 + (i * 1.12);
    shape(s, pptx.ShapeType.ellipse, 0.91, y, 0.47, 0.47, i === 0 ? C.teal : (i === 1 ? C.orange : C.amber), C.white);
    text(s, a[0], 0.91, y + 0.14, 0.47, 0.13, { fontSize: 8, color: C.white, bold: true, align: 'center' });
    text(s, a[1], 1.64, y - 0.02, 4.85, 0.22, { fontSize: 13, color: C.navy, bold: true });
    text(s, a[2], 1.64, y + 0.31, 4.92, 0.22, { fontSize: 9.7, color: C.gray });
    if (i < 2) line(s, 1.64, y + 0.75, 6.95, y + 0.75, C.lightGray, 0.6);
  });
  shape(s, pptx.ShapeType.roundRect, 8.02, 1.77, 4.73, 4.85, C.navy, C.navy, 0.12);
  text(s, 'ONDE MUDAR A MARCHA', 8.36, 2.08, 3.4, 0.16, { fontSize: 8, color: '92C9C4', bold: true, charSpacing: 1.0 });
  const conditions = [
    ['SE VOLUME REAGE', 'Acelerar premiumização, distribuição e inovação.', '87C7C2'],
    ['SE CUSTO SUPERA PREÇO', 'Cortar complexidade, ampliar produtividade e rever mix.', 'F1B44C'],
    ['SE CHOQUE SE CONFIRMA', 'Adiar capex discricionário e reforçar liquidez.', 'F08A4B']
  ];
  conditions.forEach((c, i) => {
    const y = 2.66 + (i * 1.05);
    shape(s, pptx.ShapeType.roundRect, 8.36, y, 3.99, 0.76, '173A54', '173A54', 0.08);
    text(s, c[0], 8.60, y + 0.12, 1.65, 0.13, { fontSize: 7.1, color: c[2], bold: true, charSpacing: 0.7 });
    text(s, c[1], 8.60, y + 0.34, 3.35, 0.18, { fontSize: 9.4, color: C.white, bold: true });
  });
  text(s, 'Leitura crítica: o hedge cambial e o caixa diminuem vulnerabilidade financeira, mas não substituem execução comercial.', 8.36, 5.98, 3.75, 0.30, { fontSize: 8.5, color: 'D7E2E8', italic: true, valign: 'top' });
  footer(s, 'Fontes: Ambev, Release 3T26; Form 20-F 2025. Recomendações: análise do grupo.', expandScenarios ? 6 : 4);
}

// Slide 5
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Painel de monitoramento', 'Sinais simples para saber quando cada cenário está se materializando.', 'A disciplina é ligar cada indicador a uma decisão, não acompanhar números isolados.');
  const rows = [
    ['DEMANDA', 'Volume Cerveja Brasil e NAB', 'Volume cresce; consumo fora do lar melhora', 'Volume estagna; preço sustenta receita', 'Queda simultânea de volume', C.teal],
    ['MARGEM', 'Receita/hl versus CPV/hl', 'CPV/hl abaixo de receita/hl', 'Diferença estreita ou negativa', 'CPV/hl acelera rapidamente', C.orange],
    ['MACRO & RISCO', 'BRL/USD, Brent/diesel, tributação', 'Câmbio e energia cedem', 'Volatilidade e regras indefinidas', 'Desvalorização, petróleo e tributo pressionam', C.amber]
  ];
  text(s, 'EIXO', 0.72, 1.78, 1.1, 0.15, { fontSize: 7.5, color: C.gray, bold: true, charSpacing: 0.8 });
  text(s, 'INDICADOR', 2.07, 1.78, 2.1, 0.15, { fontSize: 7.5, color: C.gray, bold: true, charSpacing: 0.8 });
  text(s, 'OTIMISTA', 4.69, 1.78, 2.1, 0.15, { fontSize: 7.5, color: C.green, bold: true, charSpacing: 0.8 });
  text(s, 'BASE', 7.31, 1.78, 2.1, 0.15, { fontSize: 7.5, color: C.amber, bold: true, charSpacing: 0.8 });
  text(s, 'PESSIMISTA', 9.93, 1.78, 2.2, 0.15, { fontSize: 7.5, color: C.red, bold: true, charSpacing: 0.8 });
  rows.forEach((r, i) => {
    const y = 2.12 + (i * 1.18);
    shape(s, pptx.ShapeType.roundRect, 0.58, y, 12.17, 0.91, C.white, C.white, 0.08);
    shape(s, pptx.ShapeType.rect, 0.58, y, 0.09, 0.91, r[5], r[5]);
    text(s, r[0], 0.82, y + 0.23, 1.06, 0.21, { fontSize: 9.5, color: C.navy, bold: true });
    text(s, r[1], 2.07, y + 0.18, 2.19, 0.30, { fontSize: 10.5, color: C.ink, bold: true, valign: 'top' });
    text(s, r[2], 4.69, y + 0.18, 2.1, 0.34, { fontSize: 9.1, color: C.ink, valign: 'top' });
    text(s, r[3], 7.31, y + 0.18, 2.1, 0.34, { fontSize: 9.1, color: C.ink, valign: 'top' });
    text(s, r[4], 9.93, y + 0.18, 2.38, 0.34, { fontSize: 9.1, color: C.ink, valign: 'top' });
  });
  shape(s, pptx.ShapeType.roundRect, 0.58, 5.95, 12.17, 0.65, C.navy, C.navy, 0.10);
  text(s, 'CONCLUSÃO', 0.90, 6.16, 1.20, 0.14, { fontSize: 7.6, color: '92C9C4', bold: true, charSpacing: 0.8 });
  text(s, 'Cenário-base é o mais aderente aos dados atuais: proteger volume e margem agora; acelerar somente quando os sinais de demanda sustentarem a decisão.', 2.12, 6.09, 9.95, 0.25, { fontSize: 11, color: C.white, bold: true });
  footer(s, 'Fontes: Ambev, Release 3T26 e Form 20-F 2025; MF/SPE, Boletim Macrofiscal; BCB Focus. Síntese: grupo.', expandScenarios ? 7 : 5);
}

// Appendix
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  header(s, 'Anexo obrigatório', 'Uso de IA e checagem crítica das fontes', 'Este anexo está fora do limite de cinco slides de conteúdo.');
  const blocks = [
    ['FERRAMENTA E USO', 'ChatGPT/Codex: localização de documentos de RI, organização da análise, estruturação dos cenários e revisão.'],
    ['CHECAGEM HUMANA', 'Dados financeiros confrontados com Release 3T26 e Form 20-F 2025; macro com Boletim Macrofiscal; resultado eleitoral com TSE.'],
    ['ERRO CORRIGIDO', '“Novo governo” foi corrigido para “novo mandato”: houve reeleição presidencial; a incerteza é fiscal, regulatória e tributária.'],
    ['LIMITAÇÃO', 'IA não substitui fonte primária. O grupo deve atualizar IPCA, Focus, câmbio e divulgações corporativas na data de entrega.']
  ];
  blocks.forEach((b, i) => {
    const y = 1.80 + (i * 1.10);
    shape(s, pptx.ShapeType.roundRect, 0.78, y, 11.75, 0.82, C.white, C.lightGray, 0.08);
    shape(s, pptx.ShapeType.roundRect, 1.02, y + 0.18, 1.77, 0.42, i === 2 ? C.orange : C.teal, i === 2 ? C.orange : C.teal, 0.06);
    text(s, b[0], 1.10, y + 0.31, 1.60, 0.12, { fontSize: 6.8, color: C.white, bold: true, align: 'center', charSpacing: 0.5 });
    text(s, b[1], 3.10, y + 0.20, 8.86, 0.37, { fontSize: 10.6, color: C.ink, valign: 'top' });
  });
  text(s, 'Fontes oficiais utilizadas: RI/SEC Ambev, Ministério da Fazenda/SPE, Banco Central e TSE.', 0.78, 6.46, 8.75, 0.20, { fontSize: 8.6, color: C.gray, italic: true });
  footer(s, 'Anexo metodológico — Uso de IA permitido pelo enunciado, com responsabilidade do grupo sobre o conteúdo.', 'A', false);
}

pptx.writeFile({ fileName: outputFile });
