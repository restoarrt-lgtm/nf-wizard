export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { desc, name, dept } = req.body;

  const system = `Ты — аналитик бизнес-процессов. Верни ТОЛЬКО валидный JSON без markdown-блоков и пояснений.

Структура:
{
  "ckp": "целевой конечный продукт — 1 предложение",
  "steps": [
    { "byt": "Роль", "delat": "Действие", "imet": "Инструменты/условия" }
  ],
  "rules": [
    "Лаконичное условие или норма — одно предложение"
  ]
}

steps: хронологический порядок, условные шаги — отдельной строкой.
rules: только нормы, обязательства, ограничения, сроки — не дублировать шаги.`;

  const user = `Процесс${name ? ': ' + name : ''}${dept ? ' | Отдел: ' + dept : ''}\n\n${desc}`;

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 2000,
        system,
        messages: [{ role: 'user', content: user }]
      })
    });

    const data = await response.json();
    if (data.error) return res.status(500).json({ error: data.error.message });

    const raw = (data.content || []).map(b => b.text || '').join('');
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    res.status(200).json(parsed);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
