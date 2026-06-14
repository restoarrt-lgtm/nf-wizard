export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;
  const safeName = (name || 'процесс').toString();

  try {
    // === XLSX через XML ===
    const xmlRows = (steps || []).map((s, i) => `
      <Row ss:Index="${i + 4}">
        <Cell><Data ss:Type="Number">${i + 1}</Data></Cell>
        <Cell><Data ss:Type="String">${esc(s.byt||'')}</Data></Cell>
        <Cell><Data ss:Type="String">${esc(s.del||'')}</Data></Cell>
        <Cell><Data ss:Type="String">${esc(s.imet||'')}</Data></Cell>
      </Row>`).join('');

    const xlsxXml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="НФ">
    <Table>
      <Row ss:Index="1">
        <Cell ss:MergeAcross="3"><Data ss:Type="String">Направляющая форма — ${esc(safeName)}</Data></Cell>
      </Row>
      <Row ss:Index="2">
        <Cell ss:MergeAcross="3"><Data ss:Type="String">ЦКП: ${esc(ckp||'')}</Data></Cell>
      </Row>
      <Row ss:Index="3">
        <Cell><Data ss:Type="String">№</Data></Cell>
        <Cell><Data ss:Type="String">БЫТЬ</Data></Cell>
        <Cell><Data ss:Type="String">ДЕЛАТЬ</Data></Cell>
        <Cell><Data ss:Type="String">ИМЕТЬ</Data></Cell>
      </Row>
      ${xmlRows}
    </Table>
  </Worksheet>
</Workbook>`;

    // === DOCX через XML ===
    const makeDocx = (title, subject, paragraphs) => {
      const paras = paragraphs.map(p => {
        if (p.type === 'h1') return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="48"/></w:rPr><w:t>${esc(p.text)}</w:t></w:r></w:p>`;
        if (p.type === 'h2') return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>${esc(p.text)}</w:t></w:r></w:p>`;
        if (p.type === 'h3') return `<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>${esc(p.text)}</w:t></w:r></w:p>`;
        if (p.type === 'sub') return `<w:p><w:pPr><w:ind w:left="360"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>– ${esc(p.text)}</w:t></w:r></w:p>`;
        if (p.type === 'empty') return `<w:p/>`;
        return `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${esc(p.text)}</w:t></w:r></w:p>`;
      }).join('');

      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${paras}</w:body>
</w:document>`;

      const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

      const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

      const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

      return { xml, rels, contentTypes, rootRels };
    };

    // Инструкция
    const instrParas = [
      { type: 'h1', text: 'Инструкция' },
      { type: 'h2', text: `по ${safeName.toLowerCase()}` },
      { type: 'body', text: `ЦКП: ${ckp || ''}` },
      { type: 'empty' },
    ];
    (steps || []).forEach((s, i) => {
      instrParas.push({ type: 'h3', text: `${i + 1}. ${s.del || ''}` });
      if (s.byt) instrParas.push({ type: 'sub', text: `Кто: ${s.byt}` });
      if (s.imet) instrParas.push({ type: 'sub', text: `Инструменты: ${s.imet}` });
      instrParas.push({ type: 'empty' });
    });

    // Правила
    const rulesParas = [
      { type: 'h1', text: 'Правила' },
      { type: 'h2', text: `по ${safeName.toLowerCase()}` },
      { type: 'empty' },
    ];
    (rules || []).forEach((r, i) => {
      rulesParas.push({ type: 'h3', text: `${i + 1}. ${r}` });
      rulesParas.push({ type: 'empty' });
    });

    const toBase64 = str => Buffer.from(str, 'utf8').toString('base64');

    res.status(200).json({
      files: [
        { name: `НФ_${safeName}.xls`, content: toBase64(xlsxXml), icon: '📊' },
        { name: `Инструкция_${safeName}.docx`, content: toBase64(makeDocx('', '', instrParas).xml), icon: '📄' },
        { name: `Правила_${safeName}.docx`, content: toBase64(makeDocx('', '', rulesParas).xml), icon: '📋' }
      ]
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

function esc(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
