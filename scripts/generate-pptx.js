// Gera apresentacao.pptx com os mesmos 13 slides de apresentacao.html,
// usando texto e formas nativas do PowerPoint (editável).
// Uso: npm run pptx
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const OUT_PATH = path.join(ROOT, "apresentacao.pptx");

const pptx = new pptxgen();
pptx.defineLayout({ name: "ESCALA_VOCE", width: 13.333, height: 7.5 });
pptx.layout = "ESCALA_VOCE";
pptx.author = "Escala Você";
pptx.title = "Escala Você — Apresentação";

const W = 13.333, H = 7.5;

// ---- palette ----
const C = {
  navy: "0F172A",
  blue700: "1D4ED8",
  blue600: "2563EB",
  blue400: "60A5FA",
  blue100: "DBEAFE",
  blue50: "EFF6FF",
  indigo700: "4338CA",
  indigo50: "EEF2FF",
  slate900: "0F172A",
  slate800: "1E293B",
  slate700: "334155",
  slate600: "475569",
  slate500: "64748B",
  slate100: "F1F5F9",
  slate50: "F8FAFC",
  white: "FFFFFF",
  red500: "EF4444",
  red400: "F87171",
  red100: "FEE2E2",
  red50: "FEF2F2",
  green600: "16A34A",
  green100: "DCFCE7",
  emerald600: "059669",
  yellow300: "FDE047",
  orange400: "FB923C",
  amber100: "FEF3C7",
  amber700: "B45309",
};

function logoMark(slide, x, y, size) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: size, h: size, rectRadius: size * 0.28,
    fill: { color: C.blue700 }, line: { type: "none" },
  });
  slide.addText("📈", { x, y, w: size, h: size, fontSize: size * 28, align: "center", valign: "middle", color: C.white });
}

function divider(slide, x, y, color) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 0.7, h: 0.06, rectRadius: 0.03, fill: { color }, line: { type: "none" } });
}

function pageFooter(slide, n) {
  slide.addText(`Escala Você · ${n}`, {
    x: W - 2.6, y: H - 0.55, w: 2.3, h: 0.3, fontSize: 10, align: "right", color: "94A3B8",
  });
}

function card(slide, opts) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: opts.x, y: opts.y, w: opts.w, h: opts.h, rectRadius: opts.radius || 0.12,
    fill: { color: opts.fill || C.white }, line: opts.line || { color: C.slate100, width: 1 },
    shadow: opts.shadow === false ? undefined : { type: "outer", color: "000000", opacity: 0.08, blur: 8, offset: 3, angle: 90 },
  });
}

async function toBase64Png(filePath) {
  const buffer = await sharp(filePath).png().toBuffer();
  return "image/png;base64," + buffer.toString("base64");
}

async function main() {
  const imgMarina = await toBase64Png(path.join(ROOT, "images/perfil-marina.webp"));
  const imgJose = await toBase64Png(path.join(ROOT, "images/perfil-jose.webp"));

  // ===================================================================
  // SLIDE 1 — CAPA
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.navy };

    logoMark(s, 0.6, 0.5, 0.5);
    s.addText([
      { text: "Escala ", options: { color: C.white, bold: true } },
      { text: "Você", options: { color: C.blue400, bold: true } },
    ], { x: 1.25, y: 0.5, w: 4, h: 0.5, fontSize: 20 });

    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: 2.0, w: 5.6, h: 0.5, rectRadius: 0.25,
      fill: { color: "1E3A8A", transparency: 40 }, line: { color: C.blue400, width: 0.75 },
    });
    s.addText("🚀  PROGRAMA DE DESENVOLVIMENTO PARA EMPREENDEDORES", {
      x: 0.75, y: 2.0, w: 5.5, h: 0.5, fontSize: 11, bold: true, color: C.blue400, valign: "middle", charSpacing: 1,
    });

    s.addText([
      { text: "Você quer construir um negócio que trabalha a favor da sua vida… ", options: { color: C.white } },
      { text: "ou uma vida que trabalha para o seu negócio?", options: { color: C.blue400 } },
    ], { x: 0.6, y: 2.65, w: 11.8, h: 1.9, fontSize: 32, bold: true, valign: "top", lineSpacing: 36 });

    s.addText(
      "Seja para viabilizar uma ideia, tracionar sua startup ou estruturar o crescimento da sua empresa. " +
      "Em 8 semanas, desenvolva a clareza, a liderança e a estratégia (OKRs) necessárias para conduzir a " +
      "próxima fase do seu negócio — sem deixar a sua vida para trás.",
      { x: 0.6, y: 4.65, w: 10.8, h: 1.1, fontSize: 13.5, color: "CBD5E1", lineSpacing: 19 }
    );

    const stats = [["8", "Semanas"], ["7", "Encontros"], ["2", "Módulos"]];
    stats.forEach((st, i) => {
      const x = 0.6 + i * 1.7;
      s.addText(st[0], { x, y: 6.15, w: 1.4, h: 0.55, fontSize: 26, bold: true, color: C.white });
      s.addText(st[1].toUpperCase(), { x, y: 6.7, w: 1.6, h: 0.3, fontSize: 9, bold: true, color: C.blue400, charSpacing: 1 });
      if (i > 0) s.addShape(pptx.ShapeType.line, { x: x - 0.25, y: 6.15, w: 0, h: 0.85, line: { color: "334155", width: 1 } });
    });
  }

  // ===================================================================
  // SLIDE 2 — SOBRE O ESCALA VOCÊ
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addText("FUNDAMENTO", { x: 0.8, y: 0.55, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.blue600, charSpacing: 2 });
    s.addText("Sobre o Escala Você", { x: 0.8, y: 0.85, w: 10, h: 0.7, fontSize: 34, bold: true, color: C.slate900 });
    divider(s, 0.8, 1.55, C.blue600);

    s.addText('"O negócio é reflexo de quem o empreendedor é."', {
      x: 0.8, y: 1.9, w: 10, h: 0.6, fontSize: 22, bold: true, color: C.slate900,
    });

    s.addText([
      { text: "O Escala Você parte de uma tese simples: quando o empreendedor cresce, o negócio cresce junto. " },
      { text: "Por isso o foco do programa não é apenas a empresa — é desenvolver a pessoa que está construindo a empresa", options: { bold: true, color: C.slate900 } },
      { text: ", trazendo clareza, foco e capacidades de liderança para conduzir a próxima fase do negócio.", options: { color: C.slate600 } },
    ], { x: 0.8, y: 2.65, w: 10.3, h: 1.15, fontSize: 14, color: C.slate600, lineSpacing: 20 });

    s.addText(
      "Um processo estruturado e individual, fundamentado em metodologias reconhecidas internacionalmente, " +
      "para avaliar opções, tomar decisões melhores e traçar um plano estratégico e viável para o seu momento.",
      { x: 0.8, y: 3.85, w: 10.3, h: 0.85, fontSize: 14, color: C.slate600, lineSpacing: 20 }
    );

    const methods = [
      ["🏅", "Gallup CliftonStrengths"], ["❤", "Barrett Values"], ["🎯", "OKRs"], ["🧠", "Positive Intelligence"],
    ];
    methods.forEach((m, i) => {
      const x = 0.8 + i * 2.5;
      card(s, { x, y: 4.95, w: 2.3, h: 1.15, fill: C.slate50, line: { color: "E2E8F0", width: 1 } });
      s.addText(m[0], { x, y: 5.05, w: 2.3, h: 0.5, fontSize: 20, align: "center" });
      s.addText(m[1], { x: x + 0.1, y: 5.6, w: 2.1, h: 0.4, fontSize: 11, bold: true, color: C.slate800, align: "center" });
    });

    pageFooter(s, "02");
  }

  // ===================================================================
  // SLIDE 3 — VOCÊ JÁ SE PEGOU PENSANDO
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.slate50 };
    s.addText("O ESCALA VOCÊ É PARA VOCÊ", { x: 0, y: 0.35, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("Você já se pegou pensando:", { x: 0, y: 0.65, w: W, h: 0.55, fontSize: 28, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.25, C.blue600);

    const pains = [
      ["❓", "“Não sei se estou no caminho certo.”", C.red50],
      ["🧠", "“Tenho muitas ideias e não consigo decidir qual seguir.”", "FFF7ED"],
      ["⏰", "“Trabalho o dia inteiro e parece que não produzi nada.”", "FEFCE8"],
      ["🔥", "“Quero crescer mas fico todo tempo apagando incêndios.”", C.red50],
      ["🔗", "“Meu negócio depende demais de mim.”", C.red50],
      ["🤝", "“Tenho dificuldade para delegar.”", "FFF7ED"],
      ["🪧", "“Não sei exatamente qual deve ser meu próximo passo.”", "FEFCE8"],
      ["🔋", "“O negócio está consumindo minha vida.”", C.red50],
      ["🏃", "“Sei o que preciso fazer, mas não consigo manter a disciplina.”", "FFF7ED"],
    ];
    const cols = 3, cw = 3.95, ch = 0.95, gap = 0.2, startX = (W - (cols * cw + (cols - 1) * gap)) / 2, startY = 1.55;
    pains.forEach((p, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const x = startX + col * (cw + gap), y = startY + row * (ch + gap);
      card(s, { x, y, w: cw, h: ch, fill: C.white, line: { color: "E2E8F0", width: 1 } });
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.18, y: y + 0.22, w: 0.5, h: 0.5, rectRadius: 0.12, fill: { color: p[2] }, line: { type: "none" } });
      s.addText(p[0], { x: x + 0.18, y: y + 0.22, w: 0.5, h: 0.5, fontSize: 16, align: "center", valign: "middle" });
      s.addText(p[1], { x: x + 0.8, y: y + 0.1, w: cw - 0.95, h: ch - 0.2, fontSize: 10.5, color: C.slate700, valign: "middle" });
    });

    const calloutY = startY + 3 * (ch + gap) + 0.1;
    card(s, { x: startX, y: calloutY, w: cols * cw + (cols - 1) * gap, h: 1.05, fill: C.blue700, line: { type: "none" } });
    s.addText([
      { text: "Se você respondeu ", options: { color: C.white } },
      { text: '"sim"', options: { color: C.yellow300, bold: true } },
      { text: " para algumas dessas perguntas, provavelmente o desafio não está apenas na empresa. Está na forma como você está ", options: { color: C.white } },
      { text: "liderando essa fase do negócio.", options: { color: C.yellow300, bold: true } },
    ], { x: startX + 0.4, y: calloutY + 0.1, w: cols * cw + (cols - 1) * gap - 0.8, h: 0.85, fontSize: 12.5, align: "center", valign: "middle", lineSpacing: 17 });

    pageFooter(s, "03");
  }

  // ===================================================================
  // SLIDE 4 — O QUE VOCÊ VAI CONQUISTAR
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addText("RESULTADOS", { x: 0, y: 0.35, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("O que você vai conquistar?", { x: 0, y: 0.65, w: W, h: 0.55, fontSize: 28, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.25, C.blue600);

    const quads = [
      { title: "CLAREZA", color: C.blue700, fill: C.blue50, bullets: ["Visão clara da direção do negócio e do seu papel como líder", "Avaliação de opções e definição de prioridades reais"] },
      { title: "FOCO", color: C.indigo700, fill: C.indigo50, bullets: ["Alinhamento de objetivos de curto, médio e longo prazo", "Energia direcionada ao que realmente importa para crescer"] },
      { title: "CONFIANÇA E LIDERANÇA", color: "047857", fill: "ECFDF5", bullets: ["Liderança mais consciente para inspirar e engajar a equipe", "Capacidade real de delegar e confiar"] },
      { title: "AÇÃO", color: C.slate700, fill: C.slate100, bullets: ["Plano estratégico de 90 dias pronto para execução", "Ferramentas práticas para gerir o negócio no dia a dia"] },
    ];
    const qw = 5.2, qh = 1.9, gapx = 0.3, gapy = 0.3;
    const startX = (W - (2 * qw + gapx)) / 2, startY = 1.65;
    quads.forEach((q, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = startX + col * (qw + gapx), y = startY + row * (qh + gapy);
      card(s, { x, y, w: qw, h: qh, fill: q.fill, line: { type: "none" }, shadow: false });
      s.addText(q.title, { x: x + 0.3, y: y + 0.2, w: qw - 0.6, h: 0.35, fontSize: 14, bold: true, color: q.color });
      s.addText(q.bullets.map((b) => ({ text: b, options: { bullet: { code: "2022" }, breakLine: true } })), {
        x: x + 0.3, y: y + 0.65, w: qw - 0.6, h: qh - 0.85, fontSize: 12, color: C.slate700, lineSpacing: 16,
      });
    });

    s.addShape(pptx.ShapeType.ellipse, {
      x: W / 2 - 0.6, y: startY + qh + gapy / 2 - 0.6, w: 1.2, h: 1.2,
      fill: { color: C.white }, line: { color: C.blue100, width: 3 },
      shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 10, offset: 3, angle: 90 },
    });
    logoMark(s, W / 2 - 0.3, startY + qh + gapy / 2 - 0.3, 0.6);

    s.addText([
      { text: "Resultados de longo prazo: ", options: { bold: true, color: C.slate800 } },
      { text: "mais clareza, negócio alinhado aos seus objetivos de vida, liderança mais forte e crescimento sustentável.", options: { color: C.slate500 } },
    ], { x: 1.5, y: 6.7, w: W - 3, h: 0.5, fontSize: 12, align: "center" });

    pageFooter(s, "04");
  }

  // ===================================================================
  // SLIDE 5 — A JORNADA ESCALA VOCÊ
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.slate50 };
    s.addText("PROGRAMA", { x: 0, y: 0.4, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("A Jornada Escala Você", { x: 0, y: 0.7, w: W, h: 0.55, fontSize: 28, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.3, C.blue600);

    const mods = [
      { n: "1", color: C.blue700, tag: "MÓDULO 1", title: "Quem sou como empreendedor?", desc: "Que tipo de negócio eu quero construir para sustentar a vida que eu quero viver?" },
      { n: "2", color: C.indigo700, tag: "MÓDULO 2", title: "Minha estratégia como empreendedor", desc: "Como transformo essa visão em prioridades e metas concretas?" },
    ];
    const cw = 5.6, ch = 3.7, gap = 0.5, startX = (W - (2 * cw + gap)) / 2, startY = 2.0;
    mods.forEach((m, i) => {
      const x = startX + i * (cw + gap);
      card(s, { x, y: startY, w: cw, h: ch, fill: C.white, line: { type: "none" } });
      s.addShape(pptx.ShapeType.roundRect, { x, y: startY, w: cw, h: 0.12, fill: { color: m.color }, line: { type: "none" } });
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.4, y: startY + 0.4, w: 0.8, h: 0.8, rectRadius: 0.16, fill: { color: m.color }, line: { type: "none" } });
      s.addText(m.n, { x: x + 0.4, y: startY + 0.4, w: 0.8, h: 0.8, fontSize: 24, bold: true, color: C.white, align: "center", valign: "middle" });
      s.addText(m.tag, { x: x + 0.4, y: startY + 1.4, w: cw - 0.8, h: 0.3, fontSize: 12, bold: true, color: m.color, charSpacing: 1 });
      s.addText(m.title, { x: x + 0.4, y: startY + 1.7, w: cw - 0.8, h: 0.7, fontSize: 19, bold: true, color: C.slate900 });
      s.addText(m.desc, { x: x + 0.4, y: startY + 2.4, w: cw - 0.8, h: 1.0, fontSize: 13, color: C.slate600, lineSpacing: 18 });
    });

    pageFooter(s, "05");
  }

  // ===================================================================
  // SLIDE 6 — CONTEÚDO DOS MÓDULOS
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addText("PROGRAMA", { x: 0, y: 0.3, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("Conteúdo dos módulos", { x: 0, y: 0.58, w: W, h: 0.5, fontSize: 26, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.12, C.blue600);

    const colW = 5.9, gap = 0.4, startX = (W - (2 * colW + gap)) / 2, startY = 1.5;

    s.addText("Módulo 1: Eu como empreendedor", { x: startX, y: startY, w: colW, h: 0.4, fontSize: 15, bold: true, color: C.blue700 });
    const m1 = [
      ["Diagnóstico inicial", " (Ikigai e visão futura) para transformar frustrações e incertezas em uma direção estratégica clara e intencional."],
      ["DNA empreendedor", ". Mapeia talentos naturais (Strengths) para criar um mapa de fortalezas e identificar a zona de gênio e sabotadores."],
      ["Proposta de Valor Profissional", ": integra a história de vida ao propósito, usando dores superadas como diferencial competitivo."],
      ["Alinhamento de vida e negócio", ". Garante que o modelo do negócio e os valores do fundador sustentem a vida desejada sem drenar sua energia."],
    ];
    let y1 = startY + 0.45;
    m1.forEach((item, i) => {
      const h = 0.78;
      card(s, { x: startX, y: y1, w: colW, h, fill: C.blue50, line: { type: "none" }, shadow: false });
      s.addShape(pptx.ShapeType.roundRect, { x: startX + 0.15, y: y1 + h / 2 - 0.18, w: 0.36, h: 0.36, rectRadius: 0.08, fill: { color: C.blue700 }, line: { type: "none" } });
      s.addText(String(i + 1), { x: startX + 0.15, y: y1 + h / 2 - 0.18, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });
      s.addText([
        { text: item[0], options: { bold: true, color: C.slate900 } },
        { text: item[1], options: { color: C.slate700 } },
      ], { x: startX + 0.6, y: y1 + 0.05, w: colW - 0.75, h: h - 0.1, fontSize: 10.5, valign: "middle", lineSpacing: 13 });
      y1 += h + 0.12;
    });

    const startX2 = startX + colW + gap;
    s.addText("Módulo 2: Minha estratégia como empreendedor (OKRs)", { x: startX2, y: startY, w: colW, h: 0.6, fontSize: 15, bold: true, color: C.indigo700 });
    const m2 = [
      ["OKRs pessoais", ". Faz um mapeamento financeiro e de tempo disponível para estabelecer as primeiras hipóteses de metas do negócio."],
      ["Teste de viabilidade", ". Aplica um choque de realidade numérico para refinar e validar se as metas financeiras e de tempo estipuladas são viáveis."],
      ["Fechamento dos OKRs", ". Combate distrações, prioriza projetos de alto retorno (Zona Mágica) e consolida um plano prático focado na execução."],
    ];
    let y2 = startY + 0.7;
    m2.forEach((item, i) => {
      const h = 0.78;
      card(s, { x: startX2, y: y2, w: colW, h, fill: C.indigo50, line: { type: "none" }, shadow: false });
      s.addShape(pptx.ShapeType.roundRect, { x: startX2 + 0.15, y: y2 + h / 2 - 0.18, w: 0.36, h: 0.36, rectRadius: 0.08, fill: { color: C.indigo700 }, line: { type: "none" } });
      s.addText(String(i + 1), { x: startX2 + 0.15, y: y2 + h / 2 - 0.18, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });
      s.addText([
        { text: item[0], options: { bold: true, color: C.slate900 } },
        { text: item[1], options: { color: C.slate700 } },
      ], { x: startX2 + 0.6, y: y2 + 0.05, w: colW - 0.75, h: h - 0.1, fontSize: 10.5, valign: "middle", lineSpacing: 13 });
      y2 += h + 0.12;
    });

    pageFooter(s, "06");
  }

  // ===================================================================
  // SLIDE 7 & 8 — COMO FUNCIONA (INDIVIDUAL / GRUPO)
  // ===================================================================
  function pricingSlide({ pageNum, eyebrowText, title, badge, statValues, mods, note }) {
    const s = pptx.addSlide();
    s.background = { color: badge || !statValues ? C.slate50 : C.white };
    s.addText(eyebrowText.toUpperCase(), { x: 0, y: 0.3, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText(title, { x: 0, y: 0.58, w: W, h: 0.5, fontSize: 26, bold: true, color: C.slate900, align: "center" });
    let afterHeaderY = 1.12;
    divider(s, W / 2 - 0.35, afterHeaderY, C.blue600);
    if (badge) {
      s.addShape(pptx.ShapeType.roundRect, { x: W / 2 - 2.3, y: afterHeaderY + 0.25, w: 4.6, h: 0.4, rectRadius: 0.2, fill: { color: C.amber100 }, line: { type: "none" } });
      s.addText("🎁  " + badge, { x: W / 2 - 2.3, y: afterHeaderY + 0.25, w: 4.6, h: 0.4, fontSize: 11, bold: true, color: C.amber700, align: "center", valign: "middle" });
      afterHeaderY += 0.55;
    }

    if (statValues) {
      const sw = 3.7, sgap = 0.25, sx = (W - (3 * sw + 2 * sgap)) / 2, sy = 1.45;
      statValues.forEach((st, i) => {
        const x = sx + i * (sw + sgap);
        card(s, { x, y: sy, w: sw, h: 0.7, fill: C.blue700, line: { type: "none" }, shadow: false });
        s.addText(st[0], { x: x + 0.2, y: sy + 0.08, w: sw - 0.4, h: 0.25, fontSize: 9.5, color: "BFDBFE" });
        s.addText(st[1], { x: x + 0.2, y: sy + 0.32, w: sw - 0.4, h: 0.32, fontSize: 15, bold: true, color: C.white });
      });
    }

    const cw = 3.7, cgap = 0.25, cx = (W - (3 * cw + 2 * cgap)) / 2;
    const cy = statValues ? 2.35 : afterHeaderY + 0.25;
    const ch = 3.45;
    mods.forEach((m, i) => {
      const x = cx + i * (cw + cgap);
      card(s, { x, y: cy, w: cw, h: ch, fill: m.highlight ? C.blue50 : C.white, line: m.highlight ? { color: C.blue600, width: 1.5 } : { color: C.slate100, width: 1 } });
      if (m.priceBadge) {
        s.addShape(pptx.ShapeType.roundRect, { x: x + cw - 1.85, y: cy - 0.18, w: 1.7, h: 0.36, rectRadius: 0.18, fill: { color: C.blue700 }, line: { type: "none" } });
        s.addText(m.priceBadge, { x: x + cw - 1.85, y: cy - 0.18, w: 1.7, h: 0.36, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle" });
      }
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.25, y: cy + 0.25, w: 0.55, h: 0.55, rectRadius: 0.12, fill: { color: m.iconBg }, line: { type: "none" } });
      s.addText(m.icon, { x: x + 0.25, y: cy + 0.25, w: 0.55, h: 0.55, fontSize: 16, align: "center", valign: "middle" });
      s.addText(m.tag, { x: x + 0.25, y: cy + 0.9, w: cw - 0.5, h: 0.25, fontSize: 9.5, bold: true, color: m.tagColor, charSpacing: 1 });
      s.addText(m.title, { x: x + 0.25, y: cy + 1.15, w: cw - 0.5, h: 0.7, fontSize: 14.5, bold: true, color: C.slate900 });
      s.addText(m.meta.map((t) => ({ text: t, options: { bullet: { code: "2022" }, breakLine: true } })), {
        x: x + 0.25, y: cy + 1.85, w: cw - 0.5, h: 0.75, fontSize: 10.5, color: C.slate600, lineSpacing: 15,
      });
      card(s, { x: x + 0.25, y: cy + ch - 1.0, w: cw - 0.5, h: 0.8, fill: m.highlight ? C.white : C.slate50, line: { type: "none" }, shadow: false });
      s.addText("Investimento" + (m.perPerson ? " por pessoa" : ""), { x: x + 0.4, y: cy + ch - 0.92, w: cw - 0.8, h: 0.22, fontSize: 9, color: C.slate500 });
      s.addText(m.price, { x: x + 0.4, y: cy + ch - 0.72, w: cw - 0.8, h: 0.4, fontSize: 20, bold: true, color: m.tagColor });
      if (m.savings) {
        s.addText(m.savings, { x: x + 0.4, y: cy + ch - 0.32, w: cw - 0.8, h: 0.25, fontSize: 9, bold: true, color: C.emerald600 });
      }
    });

    if (note) {
      s.addText(note, { x: cx, y: cy + ch + 0.15, w: 3 * cw + 2 * cgap, h: 0.4, fontSize: 10, italic: true, color: C.slate500, align: "center" });
    }

    pageFooter(s, pageNum);
  }

  pricingSlide({
    pageNum: "07",
    eyebrowText: "Como funciona",
    title: "Formato individual",
    statValues: [
      ["Duração total", "8 semanas"],
      ["Total de encontros", "7 encontros"],
      ["Duração dos encontros", "1,5h cada"],
    ],
    mods: [
      { icon: "👤", iconBg: C.blue50, tag: "MÓDULO 1", tagColor: C.blue700, title: "Quem sou como empreendedor?", meta: ["4 encontros · 3 assessments", "Duração: 4 semanas"], price: "R$ 3.800" },
      { icon: "🎯", iconBg: C.indigo50, tag: "MÓDULO 2", tagColor: C.indigo700, title: "Minha estratégia como empreendedor", meta: ["3 encontros", "Duração: 4 semanas"], price: "R$ 2.800" },
      { icon: "⭐", iconBg: C.blue600, tag: "MÓDULO 1 + MÓDULO 2", tagColor: C.blue700, title: "Programa Completo", meta: ["7 encontros", "Duração: 8 semanas"], price: "R$ 6.000", savings: "✨ Economia de R$ 600", highlight: true, priceBadge: "MELHOR ESCOLHA" },
    ],
    note: "Encontros individuais 1:1 · Exercícios práticos entre os encontros · Apoio de IA para continuidade · Acompanhamento contínuo",
  });

  pricingSlide({
    pageNum: "08",
    eyebrowText: "Como funciona",
    title: "Formato em grupo até 3 pessoas",
    badge: 'Proposta especial para participantes do evento "Da ideia à Ação"',
    statValues: null,
    mods: [
      { icon: "👤", iconBg: C.blue50, tag: "MÓDULO 1", tagColor: C.blue700, title: "Quem sou como empreendedor? *", meta: ["4 encontros", "Duração: 4 semanas"], price: "R$ 1.350", perPerson: true },
      { icon: "🎯", iconBg: C.indigo50, tag: "MÓDULO 2", tagColor: C.indigo700, title: "Minha estratégia como empreendedor", meta: ["3 encontros", "Duração: 4 semanas"], price: "R$ 1.000", perPerson: true },
      { icon: "⭐", iconBg: C.blue600, tag: "MÓDULO 1 + MÓDULO 2", tagColor: C.blue700, title: "Programa Completo *", meta: ["7 encontros", "Duração: 8 semanas"], price: "R$ 2.000", perPerson: true, savings: "✨ Economia de R$ 350", highlight: true, priceBadge: "MELHOR ESCOLHA" },
    ],
    note: "* No formato em grupo, o Módulo 1 não prevê a realização do assessment individual Gallup StrengthsFinder.",
  });

  // ===================================================================
  // SLIDE 9 — POR QUE FUNCIONA
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.slate50 };
    s.addText("POR QUE ESCOLHER", { x: 0, y: 0.3, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("Por que funciona?", { x: 0, y: 0.58, w: W, h: 0.5, fontSize: 26, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.12, C.blue600);

    const bw = 11.7, bx = (W - bw) / 2;
    card(s, { x: bx, y: 1.4, w: bw, h: 1.05, fill: C.indigo700, line: { type: "none" } });
    s.addText([
      { text: "Não trabalhamos apenas o negócio. Trabalhamos ", options: { color: C.white } },
      { text: "a pessoa que lidera o negócio", options: { color: C.yellow300, bold: true } },
      { text: " — o seu bem-estar e a sua forma única de fazer a diferença.", options: { color: C.white } },
    ], { x: bx + 0.4, y: 1.55, w: bw - 0.8, h: 0.75, fontSize: 14, valign: "middle", lineSpacing: 18 });

    const pillars = [
      ["1.", "Abordagem individual e personalizada", "Encontros 1:1, sem turmas — atenção total ao seu contexto e ritmo, com exercícios práticos entre os encontros e apoio de IA para dar continuidade à reflexão."],
      ["2.", "Metodologias reconhecidas", "Gallup CliftonStrengths, Barrett Values, OKRs e Positive Intelligence, aplicados de forma prática ao seu momento."],
      ["3.", "Facilitadores experientes", "Marina Chekmysheva (+450 mentorados) e José Almeida (16 anos escalando a BuscaOnibus a 4 milhões de usuários/mês) conduzem o programa lado a lado com você."],
      ["4.", "Resultado prático", "O programa termina com um plano estratégico de 90 dias, pronto para ser executado — não fica só na reflexão."],
    ];
    const pw = (bw - 0.3) / 2, ph = 1.55, pgap = 0.3, pStartY = 2.65;
    pillars.forEach((p, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = bx + col * (pw + 0.3), y = pStartY + row * (ph + pgap);
      card(s, { x, y, w: pw, h: ph, fill: C.white, line: { color: "E2E8F0", width: 1 } });
      s.addText([
        { text: p[0] + " ", options: { color: C.blue600, bold: true } },
        { text: p[1], options: { color: C.slate900, bold: true } },
      ], { x: x + 0.3, y: y + 0.18, w: pw - 0.6, h: 0.35, fontSize: 13.5 });
      s.addText(p[2], { x: x + 0.3, y: y + 0.55, w: pw - 0.6, h: ph - 0.7, fontSize: 10.5, color: C.slate600, lineSpacing: 14 });
    });

    pageFooter(s, "09");
  }

  // ===================================================================
  // SLIDE 10 — ANTES E DEPOIS
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addText("RESULTADOS", { x: 0, y: 0.4, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("A transformação", { x: 0, y: 0.7, w: W, h: 0.55, fontSize: 28, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.3, C.blue600);

    const cw = 5.6, gap = 0.5, startX = (W - (2 * cw + gap)) / 2, startY = 1.8, ch = 4.3;
    const antes = ["Excesso de trabalho e sobrecarga constante", "Dúvidas que não deixam o negócio avançar", "Falta de foco e dispersão de energia", "Dificuldade para liderar e inspirar a equipe", "Prioridades confusas e urgências constantes"];
    const depois = ["Clareza sobre a direção do negócio", "Foco no que realmente importa para crescer", "Liderança mais consciente e eficaz", "Capacidade real de delegar e confiar", "Empresa alinhada aos seus objetivos de vida"];

    card(s, { x: startX, y: startY, w: cw, h: ch, fill: C.white, line: { color: C.red100, width: 2 } });
    s.addText("✕  Antes", { x: startX + 0.35, y: startY + 0.3, w: cw - 0.7, h: 0.4, fontSize: 18, bold: true, color: C.red500 });
    s.addText(antes.map((a) => ({ text: a, options: { bullet: { code: "2716" }, breakLine: true } })), {
      x: startX + 0.35, y: startY + 0.9, w: cw - 0.7, h: ch - 1.1, fontSize: 13, color: C.slate700, lineSpacing: 26,
    });

    const startX2 = startX + cw + gap;
    card(s, { x: startX2, y: startY, w: cw, h: ch, fill: C.white, line: { color: C.green100, width: 2 } });
    s.addText("✓  Depois", { x: startX2 + 0.35, y: startY + 0.3, w: cw - 0.7, h: 0.4, fontSize: 18, bold: true, color: C.green600 });
    s.addText(depois.map((a) => ({ text: a, options: { bullet: { code: "2713" }, breakLine: true } })), {
      x: startX2 + 0.35, y: startY + 0.9, w: cw - 0.7, h: ch - 1.1, fontSize: 13, color: C.slate700, lineSpacing: 26,
    });

    pageFooter(s, "10");
  }

  // ===================================================================
  // SLIDE 11 — FACILITADORES
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.slate50 };
    s.addText("TIME", { x: 0, y: 0.3, w: W, h: 0.3, fontSize: 12, bold: true, color: C.blue600, align: "center", charSpacing: 2 });
    s.addText("Quem guiará você nesta jornada", { x: 0, y: 0.58, w: W, h: 0.5, fontSize: 26, bold: true, color: C.slate900, align: "center" });
    divider(s, W / 2 - 0.35, 1.12, C.blue600);

    const cw = 5.9, gap = 0.4, startX = (W - (2 * cw + gap)) / 2, startY = 1.5, ch = 4.6;

    card(s, { x: startX, y: startY, w: cw, h: ch, fill: C.white, line: { color: "E2E8F0", width: 1 } });
    s.addImage({ data: imgMarina, x: startX + 0.35, y: startY + 0.35, w: 1.1, h: 1.1, rounding: true });
    s.addText("Marina Chekmysheva", { x: startX + 1.6, y: startY + 0.4, w: cw - 1.9, h: 0.4, fontSize: 15, bold: true, color: C.slate900 });
    s.addText("Consultora internacional de desenvolvimento humano", { x: startX + 1.6, y: startY + 0.78, w: cw - 1.9, h: 0.6, fontSize: 10.5, bold: true, color: C.blue600 });
    s.addText([
      { text: "Mais de 15 anos de experiência em consultoria e gestão de pessoas, atuando no Brasil, Rússia, Índia, Chile, Argentina e México.\n\n", options: { breakLine: true } },
      { text: "Já acompanhou mais de ", options: {} },
      { text: "450 mentorados", options: { bold: true, color: C.slate900 } },
      { text: " em processos de desenvolvimento e transição de carreira.\n\n", options: { breakLine: true } },
      { text: "Criadora da ", options: {} },
      { text: "metodologia Realizar-Se", options: { bold: true, color: C.blue700 } },
      { text: ", combinando desenvolvimento humano, estratégia e IA como apoio à reflexão.", options: {} },
    ], { x: startX + 0.35, y: startY + 1.65, w: cw - 0.7, h: ch - 1.9, fontSize: 10.5, color: C.slate600, lineSpacing: 14 });

    const startX2 = startX + cw + gap;
    card(s, { x: startX2, y: startY, w: cw, h: ch, fill: C.white, line: { color: "E2E8F0", width: 1 } });
    s.addImage({ data: imgJose, x: startX2 + 0.35, y: startY + 0.35, w: 1.1, h: 1.1, rounding: true });
    s.addText("José Almeida", { x: startX2 + 1.6, y: startY + 0.4, w: cw - 1.9, h: 0.4, fontSize: 15, bold: true, color: C.slate900 });
    s.addText("Engenheiro de software internacional", { x: startX2 + 1.6, y: startY + 0.78, w: cw - 1.9, h: 0.6, fontSize: 10.5, bold: true, color: C.indigo700 });
    s.addText([
      { text: "Mais de 20 anos em tecnologia e gestão de negócios digitais em Portugal, Reino Unido e Brasil.\n\n", options: { breakLine: true } },
      { text: "Liderou por ", options: {} },
      { text: "16 anos toda a jornada da BuscaOnibus", options: { bold: true, color: C.slate900 } },
      { text: ", escalando a plataforma até ", options: {} },
      { text: "4 milhões de usuários mensais", options: { bold: true, color: C.slate900 } },
      { text: " e conduzindo o Exit para um grupo internacional.\n\n", options: { breakLine: true } },
      { text: "Especialista na implementação prática da ", options: {} },
      { text: "metodologia OKR", options: { bold: true, color: C.indigo700 } },
      { text: " para transformar prioridades em execução real.", options: {} },
    ], { x: startX2 + 0.35, y: startY + 1.65, w: cw - 0.7, h: ch - 1.9, fontSize: 10.5, color: C.slate600, lineSpacing: 14 });

    pageFooter(s, "11");
  }

  // ===================================================================
  // SLIDE 12 — CONTATO
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.navy }, line: { type: "none" } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: H - 1.05, w: W, h: 1.05, fill: { color: C.navy }, line: { type: "none" } });

    s.addShape(pptx.ShapeType.roundRect, { x: 1, y: 1.7, w: 0.8, h: 0.8, rectRadius: 0.18, fill: { color: C.blue100 }, line: { type: "none" } });
    s.addText("🤝", { x: 1, y: 1.7, w: 0.8, h: 0.8, fontSize: 22, align: "center", valign: "middle" });
    s.addText("Pronto para dar o\npróximo passo?", { x: 1, y: 2.7, w: 5, h: 1.3, fontSize: 26, bold: true, color: C.slate900, lineSpacing: 32 });
    s.addText("Se fizer sentido para você, seguimos juntos nesse processo com clareza e consciência.", {
      x: 1, y: 4.1, w: 4.6, h: 0.9, fontSize: 13, color: C.slate600, lineSpacing: 18,
    });

    const rows = [
      ["💬", "Celular / WhatsApp", "+55 (48) 99117-2456", C.green100, C.emerald600],
      ["🌐", "Website", "escala-voce.web.app", C.blue100, C.blue600],
    ];
    rows.forEach((r, i) => {
      const y = 2.35 + i * 1.15;
      s.addShape(pptx.ShapeType.roundRect, { x: 7.3, y, w: 0.8, h: 0.8, rectRadius: 0.18, fill: { color: r[3] }, line: { type: "none" } });
      s.addText(r[0], { x: 7.3, y, w: 0.8, h: 0.8, fontSize: 20, align: "center", valign: "middle" });
      s.addText(r[1], { x: 8.3, y: y + 0.05, w: 4, h: 0.3, fontSize: 11, bold: true, color: C.slate500 });
      s.addText(r[2], { x: 8.3, y: y + 0.35, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.slate900 });
    });

    pageFooter(s, "12");
  }

  // ===================================================================
  // SLIDE 13 — OBRIGADO
  // ===================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: C.navy };
    s.addShape(pptx.ShapeType.roundRect, { x: W / 2 - 0.6, y: 2.4, w: 1.2, h: 1.2, rectRadius: 0.3, fill: { color: C.blue700 }, line: { type: "none" } });
    s.addText("📈", { x: W / 2 - 0.6, y: 2.4, w: 1.2, h: 1.2, fontSize: 36, align: "center", valign: "middle" });
    s.addText("OBRIGADO", { x: 0, y: 3.8, w: W, h: 1, fontSize: 44, bold: true, color: C.white, align: "center" });
    s.addText("Escala Você — Programa de Desenvolvimento para Empreendedores", {
      x: 0, y: 4.75, w: W, h: 0.4, fontSize: 13, color: C.blue400, align: "center",
    });
  }

  await pptx.writeFile({ fileName: OUT_PATH });
  console.log("PPTX gerado em " + OUT_PATH);
}

main();
