#!/bin/bash
# 芯片齐套管理系统 - 启动脚本
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/backend"
DATA_DIR="$SCRIPT_DIR/data"
PID_FILE="$DATA_DIR/chipkit.pid"

mkdir -p "$DATA_DIR"

start() {
    if [ -f "$PID_FILE" ] && kill -0 $(cat "$PID_FILE") 2>/dev/null; then
        echo "✅ 服务已在运行 (PID: $(cat $PID_FILE))"
        return
    fi
    echo "🚀 启动芯片齐套管理系统..."
    cd "$BACKEND_DIR"
    PYTHONPATH=. nohup python3 -u server.py > "$DATA_DIR/server.log" 2>&1 &
    echo $! > "$PID_FILE"
    sleep 2
    if kill -0 $(cat "$PID_FILE") 2>/dev/null; then
        echo "✅ 服务已启动: http://localhost:8765"
        echo "   日志: $DATA_DIR/server.log"
    else
        echo "❌ 启动失败，查看日志: $DATA_DIR/server.log"
        exit 1
    fi
}

stop() {
    if [ -f "$PID_FILE" ]; then
        PID=$(cat "$PID_FILE")
        if kill -0 $PID 2>/dev/null; then
            echo "🛑 停止服务 (PID: $PID)..."
            kill $PID
            sleep 2
            if kill -0 $PID 2>/dev/null; then
                kill -9 $PID
            fi
            echo "✅ 已停止"
        fi
        rm -f "$PID_FILE"
    else
        echo "⚠️ 服务未运行"
    fi
}

status() {
    if [ -f "$PID_FILE" ] && kill -0 $(cat "$PID_FILE") 2>/dev/null; then
        echo "✅ 运行中 (PID: $(cat $PID_FILE), 端口: 8765)"
        echo "   访问: http://localhost:8765"
    else
        echo "❌ 未运行"
    fi
}

migrate() {
    echo "📥 数据迁移（从Excel导入）..."
    cd "$SCRIPT_DIR/scripts"
    PYTHONPATH="$BACKEND_DIR:$PYTHONPATH" python3 migrate_fast.py
    PYTHONPATH="$BACKEND_DIR:$PYTHONPATH" python3 migrate_inventory.py
    echo "✅ 迁移完成"
}

case "${1:-start}" in
    start)   start ;;
    stop)    stop ;;
    restart) stop; sleep 1; start ;;
    status)  status ;;
    migrate) migrate ;;
    logs)    tail -f "$DATA_DIR/server.log" ;;
    *)
        echo "用法: $0 {start|stop|restart|status|migrate|logs}"
        exit 1
        ;;
esac