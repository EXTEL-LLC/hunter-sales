#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# update_feed.sh — Генерація XML-фіду та публікація на Netlify
#
# ПЕРШИЙ ЗАПУСК (з файлом «Emos Прайс»):
#   bash update_feed.sh
#   → Скрипт сам знайде обидва файли на Робочому столі
#
# НАСТУПНІ ОНОВЛЕННЯ (тільки «Асортимент»):
#   bash update_feed.sh
#   → Ціни беруться з кешу data/price_cache.json
#
# Потребує: python3, openpyxl, netlify-cli
#   pip install openpyxl
#   npm install -g netlify-cli
#   netlify login    (один раз)
#   netlify link     (один раз — прив'язати до сайту)
# ─────────────────────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo ""
echo "════════════════════════════════════════════════"
echo "  Оновлення XML-фіду Telemart"
echo "════════════════════════════════════════════════"
echo ""

# Генерація feed.xml
python3 process_feed.py "$@"

echo ""
echo "🚀 Публікуємо на Netlify..."
netlify deploy --dir=web --prod

echo ""
echo "════════════════════════════════════════════════"
echo "  ✅ Готово! XML оновлено за постійним URL."
echo "  📎 feed.xml доступний за адресою сайту/feed.xml"
echo "════════════════════════════════════════════════"
echo ""
