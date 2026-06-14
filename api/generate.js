import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Header, AlignmentType, HeadingLevel, BorderStyle } from 'docx';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;
  const safeName = (name || 'процесс').toString();

  try {
    // === XLSX ===
    const wb = XLSX.utils.book_new();
    const wsData = [
      [`Направляющая форма — ${safeName}`],
      [`ЦКП: ${ckp || ''}`],
      ['№', 'БЫТЬ', 'ДЕЛАТЬ', 'ИМЕТЬ'],
      ...(steps || []).map((s, i) => [i + 1, s.byt || '', s.del || '', s.imet || ''])
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 5 }, { wch: 22 }, { wch: 30 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws, 'НФ');
    const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // === DOCX helpers ===
    const FONT = 'Arial', BLACK = '000000';
    const body = t => new Paragraph({
      alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0, line: 276 },
      children: [new TextRun({ text: t, font: FONT, size: 24, color: BLACK })]
    });
    const h3 = (n, t) => new Paragraph({
      heading: HeadingLevel.HEADING_3, alignment: AlignmentType.BOTH,
      spacing: { before: 140, after: 80 },
      children: [new TextRun({ text: `${n}. ${t}`, font: FONT, size: 28, bold: true, color: BLACK })]
    });
    const sub = t => new Paragraph({
      alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0, line: 276 },
      indent: { left: 360 },
      children: [new TextRun({ text: `\u2013 ${t}`, font: FONT, size: 24, color: BLACK })]
    });
    const empty = () => new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun('')] });
    const center = (t, size) => new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: t, font: FONT, size, color: BLACK })]
    });
    const makeHeader = () => new Header({
      children: [new Paragraph({
        alignment: AlignmentType.LEFT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'B4B2A9', space: 4 } },
        children: [new TextRun({ text: dept || 'Отдел', font: FONT, size: 20, color: '888780' })]
      })]
    });
    const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } };
    const styles = { default: { document: { run: { font: FONT, size: 24, color: BLACK } } } };

    // === Инструкция ===
    const instrChildren = [
      center('Инструкция', 48),
      center(`по ${safeName.toLowerCase()}`, 40),
      body(`ЦКП: ${ckp || ''}`),
      empty(),
    ];
    (steps || []).forEach((s, i) => {
      instrChildren.push(h3(i + 1, s.del || ''));
      if (s.byt) instrChildren.push(sub(`Кто: ${s.byt}`));
      if (s.imet) instrChildren.push(sub(`Инструменты: ${s.imet}`));
      instrChildren.push(empty());
    });

    const instrDoc = new Document({
      styles,
      sections: [{ properties: pageProps, headers: { default: makeHeader() }, children: instrChildren }]
    });
    const instrBuffer = await Packer.toBuffer(instrDoc);

    // === Правила ===
    const rulesChildren = [
      center('Правила', 48),
      center(`по ${safeName.toLowerCase()}`, 40),
      body(`Правила вступают в силу при выполнении процесса «${safeName}».`),
      empty(),
    ];
    (rules || []).forEach((r, i) => {
      rulesChildren.push(h3(i + 1, String(r)));
      rulesChildren.push(empty());
    });

    const rulesDoc = new Document({
      styles,
      sections: [{ properties: pageProps, headers: { default: makeHeader() }, children: rulesChildren }]
    });
    const rulesBuffer = await Packer.toBuffer(rulesDoc);

    const toBase64 = buf => Buffer.from(buf).toString('base64');

    res.status(200).json({
      files: [
        { name: `НФ_${safeName}.xlsx`, content: toBase64(xlsxBuffer), icon: '📊' },
        { name: `Инструкция_${safeName}.docx`, content: toBase64(instrBuffer), icon: '📄' },
        { name: `Правила_${safeName}.docx`, content: toBase64(rulesBuffer), icon: '📋' }
      ]
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
