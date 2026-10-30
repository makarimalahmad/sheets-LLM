#!/bin/bash
# Hentikan HANYA proses Sheets LLM (tidak mematikan program lain)
echo "Menghentikan server Sheets LLM..."
pkill -f "run_server.py" 2>/dev/null || true
sleep 1

echo "Menjalankan ulang server Sheets LLM..."
sudo PORT=80 nohup .venv/bin/python run_server.py > app.log 2>&1 &

sleep 2
echo "Status server:"
tail -n 10 app.log
