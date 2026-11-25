// Small, uncompressed ZIP writer for the six XML parts of an XLSX workbook.
// No network requests or spreadsheet runtime are required.
function zip(files: Record<string, string>) {
  const encoder = new TextEncoder()
  const local: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  let centralSize = 0
  const header = (size: number) => new Uint8Array(size)
  for (const [path, content] of Object.entries(files)) {
    const name = encoder.encode(path)
    const data = encoder.encode(content)
    let crc = 0xffffffff
    for (const byte of data) {
      crc ^= byte
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
    }
    crc = (crc ^ 0xffffffff) >>> 0
    const entry = header(30 + name.length)
    const view = new DataView(entry.buffer)
    view.setUint32(0, 0x04034b50, true)
    view.setUint16(4, 20, true)
    view.setUint32(14, crc, true)
    view.setUint32(18, data.length, true)
    view.setUint32(22, data.length, true)
    view.setUint16(26, name.length, true)
    entry.set(name, 30)
    local.push(entry, data)
    const index = header(46 + name.length)
    const iv = new DataView(index.buffer)
    iv.setUint32(0, 0x02014b50, true)
    iv.setUint16(4, 20, true)
    iv.setUint16(6, 20, true)
    iv.setUint32(16, crc, true)
    iv.setUint32(20, data.length, true)
    iv.setUint32(24, data.length, true)
    iv.setUint16(28, name.length, true)
    iv.setUint32(42, offset, true)
    index.set(name, 46)
    central.push(index)
    centralSize += index.length
    offset += entry.length + data.length
  }
  const end = header(22)
  const ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, central.length, true)
  ev.setUint16(10, central.length, true)
  ev.setUint32(12, centralSize, true)
  ev.setUint32(16, offset, true)
  return new Blob([...local, ...central, end] as BlobPart[], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  })
}

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement("a")
  link.href = url
  link.download = filename
  document.body.append(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000)
}

const xml = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const textCell = (ref: string, value: string, style = 0) =>
  `<c r="${ref}" s="${style}" t="inlineStr"><is><t>${xml(value)}</t></is></c>`
const formulaCell = (ref: string, formula: string, value: number, style = 2) =>
  `<c r="${ref}" s="${style}"><f>${xml(formula)}</f><v>${value}</v></c>`
const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const payments = ["Pix", "Dinheiro", "Cartão de débito", "Cartão de crédito", "Boleto", "Outro"]

export function downloadSalesWorkbook() {
  const headings = ["Data", "Cliente", "O que vendi", "Valor (R$)", "Como pagou", "Já recebi?"]
  const sales = `<worksheet xmlns="${ns}">
    <sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
    <cols><col min="1" max="1" width="15" customWidth="1"/><col min="2" max="3" width="29" customWidth="1"/><col min="4" max="6" width="23" customWidth="1"/></cols>
    <sheetData>
      <row r="1">${textCell("A1", "Uma linha por venda. Preencha só as colunas A a F. Apague a linha de exemplo quando começar.")}</row>
      <row r="2">${textCell("A2", "Digite valores como números; em Já recebi? use Sim ou Não. As fórmulas incluem todas as linhas a partir da 4.")}</row>
      <row r="3">${headings.map((title, index) => textCell(`${"ABCDEF"[index]}3`, title, 1)).join("")}</row>
      <row r="4">${textCell("A4", "03/11/2026")}${textCell("B4", "Maria Souza (exemplo)")}${textCell("C4", "Bolo de pote")}<c r="D4" s="2"><v>45</v></c>${textCell("E4", "Pix")}${textCell("F4", "Sim")}</row>
    </sheetData>
    <autoFilter ref="A3:F1048576"/>
    <mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="A2:F2"/></mergeCells>
    <dataValidations count="2">
      <dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Escolha uma forma de pagamento" error="Use uma opção da lista." sqref="E4:E1048576"><formula1>"${payments.join(",")}"</formula1></dataValidation>
      <dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Escolha Sim ou Não" error="Use Sim ou Não." sqref="F4:F1048576"><formula1>"Sim,Não"</formula1></dataValidation>
    </dataValidations>
  </worksheet>`
  const amounts = "Vendas!D4:D1048576"
  const received = "Vendas!F4:F1048576"
  const methods = "Vendas!E4:E1048576"
  const metrics: [string, string, number, number?][] = [
    ["Total vendido", `SUM(${amounts})`, 45],
    ["Já recebi", `SUMIF(${received},"Sim",${amounts})`, 45],
    ["Falta receber", "B4-B5", 0],
    ["Nº de vendas", `COUNT(${amounts})`, 1, 0],
    ["Ticket médio", "IFERROR(B4/B7,0)", 45],
    ...payments.map((payment): [string, string, number] => [
      `Recebido · ${payment}`,
      `SUMIFS(${amounts},${methods},"${payment}",${received},"Sim")`,
      payment === "Pix" ? 45 : 0,
    ]),
  ]
  const summary = `<worksheet xmlns="${ns}">
    <cols><col min="1" max="1" width="36" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/></cols>
    <sheetData>
      <row r="1">${textCell("A1", "Resumo automático", 1)}</row>
      <row r="2">${textCell("A2", "Não edite as fórmulas. Os valores iniciais correspondem à venda de exemplo.")}</row>
      ${metrics.map(([label, formula, cached, style], index) => `<row r="${index + 4}">${textCell(`A${index + 4}`, label)}${formulaCell(`B${index + 4}`, formula, cached, style ?? 2)}</row>`).join("")}
    </sheetData>
  </worksheet>`
  download(zip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${[1, 2].map((i) => `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Vendas" sheetId="1" r:id="rId1"/><sheet name="Resumo" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2].map((i) => `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`).join("")}<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<styleSheet xmlns="${ns}"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;R$&quot; #,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F766E"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    "xl/worksheets/sheet1.xml": sales,
    "xl/worksheets/sheet2.xml": summary,
  }), "controle-de-vendas.xlsx")
}

export function downloadOrdersForm() {
  const blankRows = Array.from({ length: 7 }, () => "<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>").join("")
  const html = `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Ficha de Controle de Pedidos e Parcelas</title>
    <style>
      @page { size: A4; margin: 15mm; }
      * { box-sizing: border-box; }
      body { font: 14px Arial, sans-serif; color: #172b29; max-width: 180mm; margin: 24px auto; padding: 0 12px; }
      h1 { font-size: 23px; color: #115e59; margin: 8px 0 24px; }
      h2 { font-size: 17px; padding: 9px; background: #eef7f5; border-bottom: 2px solid #115e59; }
      p { line-height: 2; } .fields { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin: 15px 0; }
      table { width: 100%; border-collapse: collapse; margin: 18px 0; table-layout: fixed; }
      th, td { border: 1px solid #526b67; padding: 9px; height: 34px; text-align: left; }
      th:nth-child(2) { width: 43%; } th:first-child { width: 12%; }
      .installment { padding: 14px 0; border-bottom: 1px solid #7d908d; break-inside: avoid; }
      .box { display: inline-block; width: 14px; height: 14px; border: 1px solid #172b29; vertical-align: middle; margin: 0 4px 0 12px; }
      button { min-height: 44px; background: #115e59; color: white; border: 0; border-radius: 8px; padding: 10px 18px; font-size: 14px; cursor: pointer; }
      .hint { color: #475569; font-size: 14px; }
      @media print { body { margin: 0; padding: 0; } .no-print { display: none; } h2 { print-color-adjust: exact; } }
    </style></head><body>
    <div class="no-print"><button onclick="window.print()">Imprimir / Salvar como PDF</button><p class="hint">Use papel A4. Para guardar em PDF, escolha “Salvar como PDF” na impressão.</p></div>
    <p>Conectando Oportunidades</p><h1>Ficha de Controle de Pedidos e Parcelas</h1>
    <p>Nome: __________________________________________________________________</p>
    <p>Contato: _________________________________________________________________</p>
    <p>Endereço: ________________________________________________________________</p>
    <p>Data do pedido: ____/____/________</p>
    <table><thead><tr><th>Quant.</th><th>Produto</th><th>Valor Unit.</th><th>Total</th></tr></thead>
    <tbody>${blankRows}<tr><td colspan="3"><strong>Valor Total</strong></td><td>R$ ________</td></tr></tbody></table>
    <h2>Pagamento</h2>
    <p><span class="box"></span> Cartão <span class="box"></span> Dinheiro <span class="box"></span> Pix</p>
    <div class="fields"><span>Entrada: R$ __________________</span><span>Parcela: __________ Valor: R$ __________</span></div>
    <h2>Controle de pagamento de parcelas</h2>
    ${[1, 2, 3, 4].map((n) => `<div class="installment"><strong>${n}ª parcela</strong> &nbsp; Valor: R$ __________ &nbsp; Data: ____/____/____ <span class="box"></span> Pago</div>`).join("")}
    </body></html>`
  download(new Blob([html], { type: "text/html;charset=utf-8" }), "ficha-de-pedidos-e-parcelas.html")
}

export function downloadResumeTemplate() {
  const document = `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Modelo de Currículo</title>
  <style>
    @page { size: A4; margin: 20mm; }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      padding: 20px 12px;
      background: #ffffff;
      color: #000000;
      font-family: Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.35;
    }
    .help-bar {
      display: flex;
      width: min(100%, 210mm);
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      margin: 0 auto 16px;
      padding: 12px 0;
      border-bottom: 1px solid #000000;
      background: #ffffff;
    }
    .help-bar p { margin: 0; color: #000000; }
    .help-bar button {
      flex: 0 0 auto;
      min-height: 42px;
      padding: 9px 15px;
      border: 1px solid #000000;
      border-radius: 0;
      background: #ffffff;
      color: #000000;
      font: 700 11pt Arial, sans-serif;
      cursor: pointer;
    }
    .page {
      width: min(100%, 210mm);
      min-height: 297mm;
      margin: 0 auto;
      padding: 20mm;
      background: #ffffff;
      outline: none;
    }
    header { margin-bottom: 16pt; text-align: left; }
    h1 {
      margin: 0 0 6pt;
      color: #000000;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: 700;
      line-height: 1.2;
    }
    header p { margin: 2pt 0; color: #000000; }
    section { margin-top: 14pt; page-break-inside: avoid; }
    h2 {
      margin: 0 0 7pt;
      padding: 0 0 3pt;
      border: 0;
      border-bottom: 1px solid #000000;
      color: #000000;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: 700;
      line-height: 1.2;
    }
    p { margin: 4pt 0; }
    ul { margin: 4pt 0 7pt; padding-left: 18pt; }
    li { margin: 2pt 0; }
    .example, .tip {
      color: #555555;
      font-style: italic;
    }
    .tip {
      margin-top: 7pt;
      padding: 0;
      border: 0;
      background: #ffffff;
    }
    .experience + .experience { margin-top: 9pt; }
    @media (max-width: 720px) {
      body { padding: 10px 8px; }
      .help-bar { align-items: stretch; flex-direction: column; }
      .page { min-height: 0; padding: 12mm 8mm; }
    }
    @media print {
      @page { size: A4; margin: 0; }
      body {
        margin: 0;
        padding: 15mm 20mm;
        background: #ffffff;
        font-size: 10pt;
        line-height: 1.12;
      }
      .no-print { display: none !important; }
      .page {
        width: auto;
        min-height: 0;
        margin: 0;
        padding: 0;
      }
      header { margin-bottom: 7pt; }
      h1 { margin-bottom: 4pt; }
      header p { margin: 2pt 0; }
      section { margin-top: 7pt; }
      h2 {
        margin-bottom: 4pt;
        padding-bottom: 2pt;
        font-size: 12pt;
        line-height: 1.1;
        break-after: avoid;
      }
      p { margin: 3pt 0; }
      ul { margin: 3pt 0 4pt; }
      li { margin: 1pt 0; }
      .tip { margin-top: 4pt; }
      .experience + .experience { margin-top: 5pt; }
    }
  </style>
</head>
<body>
  <div class="help-bar no-print">
    <p>Clique em qualquer texto abaixo para apagar os exemplos e digitar seus dados. Quando terminar, clique no botão ao lado.</p>
    <button type="button" onclick="window.print()">Salvar em PDF / Imprimir</button>
  </div>

  <main class="page" contenteditable="true" spellcheck="true">
    <header>
      <h1>[SEU NOME COMPLETO]</h1>
      <p>[Cidade] – [UF] | [Telefone/WhatsApp com DDD] | [E-mail profissional]</p>
      <p>[Link do LinkedIn, se tiver] | [Disponibilidade, ex: horário comercial, turnos, fins de semana]</p>
    </header>

    <section>
      <h2>Objetivo</h2>
      <p>[Cargo desejado], com foco em [área de atuação].</p>
      <p class="example">Exemplo: Auxiliar Administrativo, com foco em organização de documentos e atendimento ao público.</p>
    </section>

    <section>
      <h2>Resumo Profissional</h2>
      <p>[2 a 3 linhas: seu tempo de experiência, suas principais qualidades e o que você pode oferecer à empresa.]</p>
      <p class="example">Exemplo: Profissional com 3 anos de experiência em atendimento ao cliente, comunicativa, organizada e acostumada a trabalhar em equipe e sob metas. Busco uma oportunidade para contribuir com agilidade e responsabilidade.</p>
    </section>

    <section>
      <h2>Experiência Profissional</h2>
      <div class="experience">
        <p><strong>[Cargo] | [Nome da Empresa] | [Mês/Ano] – [Mês/Ano ou Atual]</strong></p>
        <ul><li>[Atividade 1]</li><li>[Atividade 2]</li></ul>
      </div>
      <div class="experience">
        <p><strong>[Cargo anterior] | [Nome da Empresa] | [Mês/Ano] – [Mês/Ano]</strong></p>
        <ul><li>[Atividade 1]</li><li>[Atividade 2]</li></ul>
      </div>
      <p class="tip"><strong>Dica:</strong> comece cada item com um verbo (atendi, organizei, controlei, auxiliei, operei) e use números quando puder. Exemplo: Atendi cerca de 80 clientes por dia, mantendo o caixa sempre conferido e sem diferenças. Liste da experiência mais recente para a mais antiga. Se não tem experiência formal, inclua trabalhos informais, voluntariado ou estágios.</p>
    </section>

    <section>
      <h2>Formação Acadêmica</h2>
      <p>[Curso ou nível de escolaridade] – [Nome da Instituição] | [Ano de conclusão ou Cursando]</p>
      <p class="example">Exemplo: Ensino Médio Completo – Escola Estadual X | 2019</p>
    </section>

    <section>
      <h2>Cursos Complementares</h2>
      <p>[Nome do Curso] – [Instituição] | [Carga horária] | [Ano]</p>
      <p>[Nome do Curso] – [Instituição] | [Carga horária] | [Ano]</p>
      <p class="example">Dica: coloque apenas cursos relacionados à vaga. Exemplo: Design de Sobrancelhas – ONG Nova Mulher | 10h | 2024</p>
    </section>

    <section>
      <h2>Habilidades</h2>
      <ul>
        <li>[Ferramenta ou habilidade técnica, ex: Pacote Office]</li>
        <li>[Habilidade comportamental, ex: comunicação, trabalho em equipe, organização, pontualidade]</li>
      </ul>
    </section>

    <section>
      <h2>Idiomas</h2>
      <p>[Idioma] – [Básico, Intermediário, Avançado ou Fluente]</p>
    </section>

    <section>
      <h2>Informações Adicionais</h2>
      <p>[Outras informações relevantes para a vaga]</p>
    </section>
  </main>
</body>
</html>`

  download(
    new Blob(["\ufeff", document], {
      type: "text/html;charset=utf-8",
    }),
    "Modelo_de_Curriculo.html",
  )
}
