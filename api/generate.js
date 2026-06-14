export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { steps, rules, ckp, name, dept } = req.body;

  try {
    // Формируем НФ как CSV для xlsx
    let nfCsv = '№,БЫТЬ,ДЕЛАТЬ,ИМЕТЬ\n';
    steps.forEach((s, i) => {
      const row = [i+1, s.byt, s.del, s.imet].map(v => `"${(v||'').replace(/"/g,'""')}"`).join(',');
      nfCsv += row + '\n';
    });

    // Формируем текст инструкции
    let instrText = `ИНСТРУКЦИЯ\n${name ? 'по ' + name.toLowerCase() : ''}\n\nЦКП: ${ckp}\n\n`;
    steps.forEach((s, i) => {
      instrText += `${i+1}. ${s.del}\n`;
      if (s.byt) instrText += `   Кто: ${s.byt}\n`;
      if (s.imet) instrText += `   Инструменты: ${s.imet}\n`;
      instrText += '\n';
    });

    // Формируем текст правил
    let rulesText = `ПРАВИЛА\n${name ? 'по ' + name.toLowerCase() : ''}\n\n`;
    rules.forEach((r, i) => {
      rulesText += `${i+1}. ${r}\n`;
    });

    res.status(200).json({
      files: [
        {
          name: `НФ_${name || 'процесс'}.csv`,
          content: Buffer.from('\uFEFF' + nfCsv).toString('base64'),
          icon: '📊'
        },
        {
          name: `Инструкция_${name || 'процесс'}.txt`,
          content: Buffer.from(instrText).toString('base64'),
          icon: '📄'
        },
        {
          name: `Правила_${name || 'процесс'}.txt`,
          content: Buffer.from(rulesText).toString('base64'),
          icon: '📋'
        }
      ]
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
