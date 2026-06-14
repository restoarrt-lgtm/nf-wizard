function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function makeXlsx(safeName, ckp, steps) {
  const rows = (steps||[]).map((s,i) => `
    <Row>
      <Cell><Data ss:Type="Number">${i+1}</Data></Cell>
      <Cell><Data ss:Type="String">${esc(s.byt)}</Data></Cell>
      <Cell><Data ss:Type="String">${esc(s.del)}</Data></Cell>
      <Cell><Data ss:Type="String">${esc(s.imet)}</Data></Cell>
    </Row>`).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Styles>
    <Style ss:ID="h"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#2C2C2A" ss:Pattern="Solid"/></Style>
    <Style ss:ID="s"><Font ss:Italic="1"/><Interior ss:Color="#F1EFE8" ss:Pattern="Solid"/></Style>
  </Styles>
  <Worksheet ss:Name="НФ">
    <Table>
      <Column ss:Width="30"/><Column ss:Width="120"/><Column ss:Width="160"/><Column ss:Width="160"/>
      <Row><Cell ss:MergeAcross="3" ss:StyleID="h"><Data ss:Type="String">Направляющая форма — ${esc(safeName)}</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="3" ss:StyleID="s"><Data ss:Type="String">ЦКП: ${esc(ckp)}</Data></Cell></Row>
      <Row>
        <Cell ss:StyleID="h"><Data ss:Type="String">№</Data></Cell>
        <Cell ss:StyleID="h"><Data ss:Type="String">БЫТЬ</Data></Cell>
        <Cell ss:StyleID="h"><Data ss:Type="String">ДЕЛАТЬ</Data></Cell>
        <Cell ss:StyleID="h"><Data ss:Type="String">ИМЕТЬ</Data></Cell>
      </Row>
      ${rows}
    </Table>
  </Worksheet>
</Workbook>`;
}

function buildZip(files) {
  const entries = [];
  let offset = 0;
  const centralHeaders = [];

  for (const [name, content] of files) {
    const nameBytes = Buffer.from(name, 'utf8');
    const contentBytes = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const crc = crc32(contentBytes);
    const local = Buffer.alloc(30 + nameBytes.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 6); local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(contentBytes.length, 18);
    local.writeUInt32LE(contentBytes.length, 22);
    local.writeUInt16LE(nameBytes.length, 26); local.writeUInt16LE(0, 28);
    nameBytes.copy(local, 30);
    centralHeaders.push({ nameBytes, crc, size: contentBytes.length, offset });
    entries.push(local, contentBytes);
    offset += local.length + contentBytes.length;
  }

  const centralDir = centralHeaders.map(({ nameBytes, crc, size, offset }) => {
    const cd = Buffer.alloc(46 + nameBytes.length);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(0, 10); cd.writeUInt16LE(0, 12); cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16); cd.writeUInt32LE(size, 20); cd.writeUInt32LE(size, 24);
    cd.writeUInt16LE(nameBytes.length, 28); cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34); cd.writeUInt16LE(0, 36); cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    nameBytes.copy(cd, 46);
    return cd;
  });

  const centralSize = centralDir.reduce((s, b) => s + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(0, 4); eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(centralHeaders.length, 8); eocd.writeUInt16LE(centralHeaders.length, 10);
  eocd.writeUInt32LE(centralSize, 12); eocd.writeUInt32LE(offset, 16); eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...entries, ...centralDir, eocd]);
}

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function makeDocxZip(title, subject, content, dept) {
  const paras = [];

  paras.push(`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="100"/></w:pPr><w:r><w:rPr><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr><w:t>${esc(title)}</w:t></w:r></w:p>`);
  paras.push(`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="100" w:after="180"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr><w:t>${esc(subject)}</w:t></w:r></w:p>`);

  for (const block of content) {
    if (block.type === 'intro') {
      paras.push(`<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${esc(block.text)}</w:t></w:r></w:p>`);
    } else if (block.type === 'step') {
      paras.push(`<w:p><w:pPr><w:spacing w:before="200" w:after="60"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>${esc(block.num)}. ${esc(block.title)}</w:t></w:r></w:p>`);
      paras.push(`<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:before="0" w:after="160"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${esc(block.desc)}</w:t></w:r></w:p>`);
    } else if (block.type === 'protocol') {
      paras.push(`<w:p><w:pPr><w:spacing w:before="160" w:after="0"/><w:pBdr><w:top w:val="single" w:sz="4" w:space="4" w:color="888888"/></w:pBdr></w:pPr></w:p>`);
      paras.push(`<w:p><w:pPr><w:spacing w:before="80" w:after="60"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>${esc(block.title)}</w:t></w:r></w:p>`);
      for (const ps of block.steps) {
        const prefix = block.steps.length > 1 ? `${ps.n}. ` : '';
        paras.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${esc(prefix + ps.text)}</w:t></w:r></w:p>`);
      }
    } else if (block.type === 'rule') {
      paras.push(`<w:p><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>${esc(block.num)}. ${esc(block.title)}</w:t></w:r></w:p>`);
      if (block.desc) paras.push(`<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:before="0" w:after="100"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${esc(block.desc)}</w:t></w:r></w:p>`);
    }
  }

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${paras.join('\n')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/></w:sectPr></w:body>
</w:document>`;

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

  return buildZip([
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', rootRels],
    ['word/document.xml', docXml],
    ['word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`],
  ]);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;
  const safeName = (name || 'процесс').toString();

  try {
    const instrPrompt = `Ты составляешь процессную инструкцию для сотрудника.

Правила:
- Первый блок — intro: одно предложение когда инструкция вступает в силу
- Далее шаги — только основные действия, без лишнего
- Название шага: короткое, в приказательной форме (глагол-инфинитив)
- Описание шага: 1-2 предложения, лаконично, в приказательной форме, упомяни инструменты если нужно
- Если шаг может пойти не по плану — добавь protocol сразу после него
- Протокол: короткое название + 1-3 шага (если шаг один — steps содержит один элемент)
- НЕ включай в инструкцию правила поведения (вежливость, прощание) — это идёт в правила
- НЕ дублируй шаги — объединяй похожие

Верни ТОЛЬКО валидный JSON без markdown:
{
  "instr": [
    { "type": "intro", "text": "..." },
    { "type": "step", "num": 1, "title": "...", "desc": "..." },
    { "type": "protocol", "title": "...", "steps": [{ "n": 1, "text": "..." }] },
    { "type": "step", "num": 2, "title": "...", "desc": "..." }
  ],
  "rules": [
    { "type": "rule", "num": 1, "title": "...", "desc": "..." }
  ]
}

Процесс: ${safeName}
Шаги НФ:
${(steps||[]).map((s,i) => `${i+1}. ${s.del} | Кто: ${s.byt} | Инструменты: ${s.imet}`).join('\n')}

Правила из НФ:
${(rules||[]).map((r,i) => `${i+1}. ${r}`).join('\n')}`;

    const aiResp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 3000,
        messages: [{ role: 'user', content: instrPrompt }]
      })
    });

    const aiData = await aiResp.json();
    const raw = (aiData.content||[]).map(b => b.text||'').join('');
    const parsed = JSON.parse(raw.replace(/```json|```/g,'').trim());

    const xlsxContent = makeXlsx(safeName, ckp||'', steps);
    const instrZip = makeDocxZip('Инструкция', `по ${safeName.toLowerCase()}`, parsed.instr || [], dept);
    const rulesZip = makeDocxZip('Правила', `по ${safeName.toLowerCase()}`, parsed.rules || [], dept);

    res.status(200).json({
      files: [
        { name: `НФ_${safeName}.xls`, content: Buffer.from(xlsxContent,'utf8').toString('base64'), icon: '📊' },
        { name: `Инструкция_${safeName}.docx`, content: instrZip.toString('base64'), icon: '📄' },
        { name: `Правила_${safeName}.docx`, content: rulesZip.toString('base64'), icon: '📋' }
      ]
    });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
}
