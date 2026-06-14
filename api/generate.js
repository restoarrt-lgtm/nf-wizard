export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;

  try {
    const safeName = (name || 'процесс').toString();

    // НФ как CSV
    let nfCsv = '\uFEFF№,БЫТЬ,ДЕЛАТЬ,ИМЕТЬ\n';
    (steps || []).forEach((s, i) => {
      const cols = [i+1, s.byt||'', s.del||'', s.imet||''];
      nfCsv += cols.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',') + '\n';
    });

    // Инструкция
    let instrText = `ИНСТРУКЦИЯ\nпо ${safeName.toLowerCase()}\n\nЦКП: ${ckp||''}\n\n`;
    (steps || []).forEach((s, i) => {
      instrText += `${i+1}. ${s.del||''}\n`;
      if (s.byt) instrText += `   Кто: ${s.byt}\n`;
      if (s.imet) instrText += `   Инструменты: ${s.imet}\n`;
      instrText += '\n';
    });

    // Правила
    let rulesText = `ПРАВИЛА\nпо ${safeName.toLowerCase()}\n\n`;
    (rules || []).forEach((r, i) => {
      rulesText += `${i+1}. ${String(r)}\n`;
    });

    const toBase64 = str => Buffer.from(str, 'utf8').toString('base64');

    res.status(200).json({
      files: [
        { name: `НФ_${safeName}.csv`, content: toBase64(nfCsv), icon: '📊' },
        { name: `Инструкция_${safeName}.txt`, content: toBase64(instrText), icon: '📄' },
        { name: `Правила_${safeName}.txt`, content: toBase64(rulesText), icon: '📋' }
      ]
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
