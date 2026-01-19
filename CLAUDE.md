# Инструкции для Claude

## После КАЖДОГО изменения кода ОБЯЗАТЕЛЬНО:

1. **Обновить версию** в `index.html`:
   - Найти `v2.X.X` в span с классом `bg-indigo-100`
   - Увеличить номер версии

2. **Добавить запись в changelog** в `index.html`:
   - Найти секцию `История изменений (Changelog)`
   - Добавить новый блок в НАЧАЛО списка с новой версией
   - Указать время по МСК

3. **Закоммитить и запушить**:
```bash
cd "/Users/mihailmirosnicenko/Desktop/vibe projects/wildberries-acts-generator"
git add .
git commit -m "vX.X.X: Краткое описание изменений

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>"
git push origin main
```

## Формат версий:
- **Major (X.0.0)** - большие изменения, новый функционал
- **Minor (0.X.0)** - добавление фич, заметные улучшения
- **Patch (0.0.X)** - багфиксы, мелкие правки

## Деплой:
- Vercel автоматически деплоит при push в main
- Сайт: wildberries-acts.krechet.space

## ВАЖНО:
- НЕ забывать про версию и changelog!
- НЕ делать изменения без коммита!
- ВСЕГДА пушить после коммита!
