#!/bin/bash
# Restored Mac environment (2026-10-09)
source "/Users/tuzi/个人项目/One/.restore/runtime.sh"
# Excel 数据阅览修改工具 · 本地启动器
# 双击此文件 → 自动启动本地服务并打开 http://127.0.0.1:58751/

cd "$(dirname "$0")" || exit 1
PORT=58751
URL="http://127.0.0.1:$PORT/"

# 已有实例在运行 → 直接打开页面，避免端口冲突
if curl -s --noproxy '*' -o /dev/null --max-time 2 "http://127.0.0.1:$PORT/"; then
  open "$URL"
  exit 0
fi

# 后台启动本地服务（日志写入 server.log，无需保持终端窗口）
nohup python3 -m http.server "$PORT" --bind 127.0.0.1 > server.log 2>&1 &

# 等服务就绪后打开浏览器（最多等 5 秒）
for i in 1 2 3 4 5; do
  if curl -s --noproxy '*' -o /dev/null --max-time 2 "http://127.0.0.1:$PORT/"; then
    break
  fi
  sleep 1
done

open "$URL"
