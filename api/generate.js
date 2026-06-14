import ExcelJS from 'exceljs';
import { Document, Packer, Paragraph, TextRun, Header, AlignmentType, HeadingLevel, BorderStyle } from 'docx';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;
  const safeName = (name || 'процесс').toString();

  try {
    // === XLSX ===
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('НФ');

    ws.mergeCells('A1:E1');
    ws.getCell('A1').value = `Направляющая форма — ${safeName}`;
    ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C2C2A' } };
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 32;

    ws.mergeCells('A2:E2');
    ws.getCell('A2').value = `ЦКП: ${ckp || ''}`;
    ws.getCell('A2').font = { italic: true, size: 10 };
    ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1EFE8' } };

    const headers = ['№', 'БЫТЬ', 'ДЕЛАТЬ', 'ИМЕТЬ', 'Комментарии'];
    const headerRow = ws.getRow(3);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C2C2A' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    headerRow.height = 30;

    (steps || []).forEach((s, i) => {
      const row = ws.getRow(i + 4);
      const bg = i % 2 === 0 ? 'FFFFFFFF' : 'FFF1EFE8';
      [i + 1, s.byt || '', s.del || '', s.imet || '', ''].forEach((v, ci) => {
        const cell = row.getCell(ci + 1);
        cell.value = v;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'top', wrapText: true };
        cell.font = { size: 11 };
      });
      row.height = 48;
    });

    ws.columns = [
      { width: 5 }, { width: 22 }, { width: 30 }, { width: 30 }, { width: 40 }
    ];

    const xlsxBuffer = await wb.xlsx.writeBuffer();

    // === DOCX Инструкция ===
    const FONT = 'Arial', BLACK = '000000';
    const body = t => new Paragraph({
      alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0, line: 276 },
      children: [new TextRun({ text: t, font: FONT, size: 24, color: BLACK })]
    });
    const h2 = t => new Paragraph({
      heading: HeadingLevel.HEADING_2, alignment: AlignmentType.BOTH,
      spacing: { before: 200, after: 120 },
      children: [new TextRun({ text: t, font: FONT, size: 32, bold: true, color: BLACK })]
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

    const instrChildren = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 }, children: [new TextRun({ text: 'Инструкция', font: FONT, size: 48, color: BLACK })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 180 }, children: [new TextRun({ text: `по ${safeName.toLowerCase()}`, font: FONT, size: 40, color: BLACK })] }),
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
      styles: { default: { document: { run: { font: FONT, size: 24, color: BLACK } } } },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.LEFT, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'B4B2A9', space: 4 } }, children: [new TextRun({ text: dept || 'Отдел', font: FONT, size: 20, color: '888780' })] })] }) },
        children: instrChildren
      }]
    });
    const instrBuffer = await Packer.toBuffer(instrDoc);

    // === DOCX Правила ===
    const rulesChildren = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 }, children: [new TextRun({ text: 'Правила', font: FONT, size: 48, color: BLACK })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 180 }, children: [new TextRun({ text: `по ${safeName.toLowerCase()}`, font: FONT, size: 40, color: BLACK })] }),
      body(`Правила вступают в силу при выполнении процесса «${safeName}».`),
      empty(),
    ];

    (rules || []).forEach((r, i) => {
      rulesChildren.push(h3(i + 1, r));
      rulesChildren.push(empty());
    });

    const rulesDoc = new Document({
      styles: { default: { document: { run: { font: FONT, size: 24, color: BLACK } } } },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.LEFT, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'B4B2A9', space: 4 } }, children: [new TextRun({ text: dept || 'Отдел', font: FONT, size: 20, color: '888780' })] })] }) },
        children: rulesChildren
      }]
    });
    const rulesBuffer = await Packer.toBuffer(rulesDoc);

    const toBase64 = buf => buf.toString('base64');

    res.status(200).json({
      files: [
        { name: `НФ_${safeName}.xlsx`, content: toBase64(Buffer.from(xlsxBuffer)), icon: '📊' },
        { name: `Инструкция_${safeName}.docx`, content: toBase64(instrBuffer), icon: '📄' },
        { name: `Правила_${safeName}.docx`, content: toBase64(rulesBuffer), icon: '📋' }
      ]
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
