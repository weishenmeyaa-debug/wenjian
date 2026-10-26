#!/bin/bash
# 格式检查器 — Mac 一键启动
cd "$(dirname "$0")"

# 检测 python3
if ! command -v python3 &>/dev/null; then
  echo "❌ 未找到 Python 3，请先安装："
  echo "   方法1: 打开「终端」输入 xcode-select --install"
  echo "   方法2: 去 https://www.python.org/downloads/ 下载安装"
  read -p "按回车退出..."
  exit 1
fi

# 创建虚拟环境
if [ ! -d "venv" ]; then
  echo "📦 首次运行，正在安装依赖（需要联网，大约 1 分钟）..."
  python3 -m venv venv
  source venv/bin/activate
  # 国内直连 pypi 常常超时，先用清华镜像，失败再回退官方源
  pip install -r requirements.txt -q -i https://pypi.tuna.tsinghua.edu.cn/simple \
    || pip install -r requirements.txt -q \
    || { echo "❌ 依赖安装失败。请检查网络；Mac 还需在「应用程序 → Python 3.x」文件夹里运行一次 Install Certificates.command"; read -p "按回车退出..."; exit 1; }
else
  source venv/bin/activate
fi

# 找一个可用端口
PORT=5001
while lsof -i :$PORT &>/dev/null; do
  PORT=$((PORT + 1))
done

# 启动
echo ""
echo "✅ 格式检查器启动成功，浏览器将自动打开..."
echo "   如果没有自动打开，请访问: http://localhost:$PORT"
echo "   关闭此窗口即可停止程序"
echo ""

# 延迟打开浏览器
(sleep 2 && open "http://localhost:$PORT") &

# 运行 Flask
python3 app.py --port $PORT
