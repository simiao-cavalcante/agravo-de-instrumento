/**
 * Script para gerar Agravo de Instrumento em formato .docx
 * Formatação fiel aos documentos reais da PGM-Rio
 *
 * Uso: node gerar_eds.js <input.json> <output.docx>
 *
 * Schema do input.json:
 * {
 *   "enderecamento": "EXCELENTÍSSIMO...",
 *   "processo": "0000000-00.0000.0.00.0000",
 *   "tipo_recurso": "Agravo de Instrumento",
 *   "embargante": "MUNICÍPIO DO RIO DE JANEIRO",
 *   "embargado": "NOME DA PARTE CONTRÁRIA",
 *   "abertura": "O MUNICÍPIO DO RIO DE JANEIRO...",
 *   "secoes": [
 *     {
 *       "tipo": "h1" | "h2" | "corpo" | "pedidos",
 *       "texto": "conteúdo da seção"
 *     }
 *   ],
 *   "intimacoes": "Requer sejam as intimações...",
 *   "fechamento": {
 *     "data": "01 de janeiro de 2025",
 *     "nome": "Manoel Simião Cavalcante Neto",
 *     "cargo": "Procurador do Município do Rio de Janeiro",
 *     "matricula": "10/331.957-1"
 *   }
 * }
 */

const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  HeadingLevel,
  Header,
  Footer,
  PageNumber,
} = require("docx");

const inputPath = process.argv[2];
const outputPath = process.argv[3];

if (!inputPath || !outputPath) {
  console.error("Uso: node gerar_contestacao.js <input.json> <output.docx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));

// Caminho da imagem do papel timbrado
const headerImagePath = path.join(__dirname, "assets", "image1.png");
const headerImageBuffer = fs.readFileSync(headerImagePath);

// ============================================================
// Constantes de formatação (extraídas dos documentos reais)
// ============================================================
const FONT_NAME = "Palatino Linotype";
const FONT_SIZE = 26; // 13pt em half-points (docx-js usa half-points)
const LINE_SPACING = 390; // 1.5 × 260 (13pt baseline)
const FOOTER_FONT_SIZE = 20; // 10pt para rodapé

// Recuo de primeira linha em DXA (1cm ≈ 567 DXA)
const FIRST_LINE_INDENT = 567;

// Margens em DXA (twentieths of a point, 1440 = 1 inch)
const MARGIN_TOP = 1987;    // 1.38"
const MARGIN_BOTTOM = 1411; // 0.98"
const MARGIN_LEFT = 2045;   // 1.42"
const MARGIN_RIGHT = 1469;  // 1.02"

// Página Letter
const PAGE_WIDTH = 12240;   // 8.5"
const PAGE_HEIGHT = 15840;  // 11"

// Imagem do header (em EMU: 914400 EMU = 1 inch)
const HEADER_IMG_WIDTH = 3239135;  // 3.54"
const HEADER_IMG_HEIGHT = 647700;  // 0.71"

// ============================================================
// Funções auxiliares
// ============================================================

function criarTextRun(texto, opts = {}) {
  return new TextRun({
    text: texto,
    font: FONT_NAME,
    size: FONT_SIZE,
    bold: opts.bold || false,
    italics: opts.italic || false,
  });
}

// Cria parágrafos de corpo com formatação padrão
// Suporta marcação **negrito** e *itálico* dentro do texto
function corpoParagrafo(texto, opts = {}) {
  const parts = [];
  // Regex para **negrito** e *itálico*
  const regex = /\*\*(.+?)\*\*|\*(.+?)\*/g;
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(texto)) !== null) {
    // Texto normal antes do match
    if (match.index > lastIndex) {
      parts.push(criarTextRun(texto.substring(lastIndex, match.index)));
    }
    if (match[1]) {
      // **negrito**
      parts.push(criarTextRun(match[1], { bold: true }));
    } else if (match[2]) {
      // *itálico*
      parts.push(criarTextRun(match[2], { italic: true }));
    }
    lastIndex = regex.lastIndex;
  }

  if (lastIndex < texto.length) {
    parts.push(criarTextRun(texto.substring(lastIndex)));
  }

  return new Paragraph({
    children: parts,
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    spacing: {
      line: LINE_SPACING,
      after: opts.afterSpacing !== undefined ? opts.afterSpacing : 120,
    },
    indent: opts.indent || { firstLine: FIRST_LINE_INDENT },
  });
}

// Título H1: centralizado, negrito, mesmo tamanho 13pt
function tituloH1(texto) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [criarTextRun(texto, { bold: true })],
    alignment: AlignmentType.CENTER,
    spacing: {
      before: 360,
      after: 240,
      line: LINE_SPACING,
    },
  });
}

// Título H2: alinhado à esquerda, negrito, mesmo tamanho 13pt
function tituloH2(texto) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [criarTextRun(texto, { bold: true })],
    alignment: AlignmentType.LEFT,
    spacing: {
      before: 240,
      after: 180,
      line: LINE_SPACING,
    },
  });
}

// Parágrafo vazio (espaçamento)
function espacamento(after = 120) {
  return new Paragraph({
    children: [],
    spacing: { after },
  });
}

// ============================================================
// Montar o documento
// ============================================================

const children = [];

// Endereçamento (centralizado, negrito)
children.push(
  new Paragraph({
    children: [criarTextRun(data.enderecamento, { bold: true })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 480, line: LINE_SPACING },
  })
);

// Número do processo e tipo do recurso
const processoText = data.tipo_recurso
  ? `${data.tipo_recurso} nº ${data.processo}`
  : `Processo nº ${data.processo}`;
children.push(
  new Paragraph({
    children: [criarTextRun(processoText)],
    alignment: AlignmentType.LEFT,
    spacing: { after: 240, line: LINE_SPACING },
  })
);

// Partes (flexível: aceita 'partes' array ou campos legados embargante/embargado)
const partes = data.partes || [];
if (partes.length === 0) {
  // Fallback para campos legados
  if (data.embargante) partes.push({ label: "Embargante", nome: data.embargante });
  if (data.embargado) partes.push({ label: "Embargado", nome: data.embargado });
}
for (let i = 0; i < partes.length; i++) {
  const isLast = i === partes.length - 1;
  children.push(
    new Paragraph({
      children: [
        criarTextRun(`${partes[i].label}: `, { bold: false }),
        criarTextRun(partes[i].nome, { bold: false }),
      ],
      alignment: AlignmentType.LEFT,
      spacing: { after: isLast ? 360 : 60, line: LINE_SPACING },
    })
  );
}

// Abertura
children.push(corpoParagrafo(data.abertura));
children.push(espacamento());

// Seções do conteúdo
for (const secao of data.secoes) {
  if (secao.tipo === "h1") {
    children.push(tituloH1(secao.texto));
  } else if (secao.tipo === "h2") {
    children.push(tituloH2(secao.texto));
  } else if (secao.tipo === "corpo") {
    const paragrafos = secao.texto.split("\n\n");
    for (const p of paragrafos) {
      if (p.trim()) {
        children.push(corpoParagrafo(p.trim()));
      }
    }
  } else if (secao.tipo === "pedidos") {
    const linhas = secao.texto.split("\n");
    for (const linha of linhas) {
      if (linha.trim()) {
        const indent = linha.match(/^[a-z]\)/)
          ? { left: 720 }
          : linha.match(/^-/)
            ? { left: 1080 }
            : undefined;
        children.push(corpoParagrafo(linha.trim(), { indent }));
      }
    }
  }
}

// Parágrafo de intimações institucionais (específico dos EDs)
if (data.intimacoes) {
  children.push(espacamento(240));
  children.push(corpoParagrafo(data.intimacoes));
}

// Fechamento
children.push(espacamento(240));
children.push(
  new Paragraph({
    children: [criarTextRun("Pede deferimento.")],
    alignment: AlignmentType.LEFT,
    spacing: { after: 480, line: LINE_SPACING },
  })
);
children.push(
  new Paragraph({
    children: [criarTextRun(`Rio de Janeiro, ${data.fechamento.data}.`)],
    alignment: AlignmentType.LEFT,
    spacing: { after: 480, line: LINE_SPACING },
  })
);
children.push(
  new Paragraph({
    children: [criarTextRun(data.fechamento.nome, { bold: true })],
    alignment: AlignmentType.LEFT,
    spacing: { line: LINE_SPACING },
  })
);
children.push(
  new Paragraph({
    children: [criarTextRun(data.fechamento.cargo)],
    alignment: AlignmentType.LEFT,
    spacing: { line: LINE_SPACING },
  })
);
children.push(
  new Paragraph({
    children: [criarTextRun(`Matrícula ${data.fechamento.matricula}`)],
    alignment: AlignmentType.LEFT,
    spacing: { line: LINE_SPACING },
  })
);

// ============================================================
// Criar documento com header (papel timbrado) e footer
// ============================================================

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: FONT_NAME,
          size: FONT_SIZE,
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: FONT_SIZE, bold: true, font: FONT_NAME },
        paragraph: {
          spacing: { before: 360, after: 240 },
          alignment: AlignmentType.CENTER,
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: FONT_SIZE, bold: true, font: FONT_NAME },
        paragraph: {
          spacing: { before: 240, after: 180 },
          outlineLevel: 1,
        },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: {
            width: PAGE_WIDTH,
            height: PAGE_HEIGHT,
          },
          margin: {
            top: MARGIN_TOP,
            bottom: MARGIN_BOTTOM,
            left: MARGIN_LEFT,
            right: MARGIN_RIGHT,
          },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  type: "png",
                  data: headerImageBuffer,
                  transformation: {
                    width: Math.round(HEADER_IMG_WIDTH / 9525), // EMU to pixels (approx)
                    height: Math.round(HEADER_IMG_HEIGHT / 9525),
                  },
                  altText: {
                    title: "Papel Timbrado PGM-Rio",
                    description: "Procuradoria Geral do Município do Rio de Janeiro",
                    name: "PGM-Rio Logo",
                  },
                }),
              ],
              alignment: AlignmentType.LEFT,
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  children: [PageNumber.CURRENT],
                  font: FONT_NAME,
                  size: FOOTER_FONT_SIZE,
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Agravo de Instrumento gerado com sucesso: ${outputPath}`);
});
