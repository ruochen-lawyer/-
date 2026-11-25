#!/bin/bash
# files_to_md.py 的 macOS 双击启动器
# 双击本文件 -> 把文件夹拖进终端窗口并回车 -> 转换完成自动在访达打开结果目录

cd "$(dirname "$0")" || exit 1

# 优先找一个可用的 Python 3.9+
PY=""
for c in python3 python /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$c" >/dev/null 2>&1; then
    PY="$c"
    break
  fi
done

if [ -t 1 ]; then clear 2>/dev/null || true; fi   # 仅在真实终端里清屏，避免重定向时输出控制字符
echo "============================================================"
echo "   文件夹批量转 Markdown  ·  文件名自动加当天日期"
echo "============================================================"
echo

if [ -z "$PY" ] || ! "$PY" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; then
  echo "未找到可用的 Python 3.9+，请先安装："
  echo "  brew install python3    或访问  https://www.python.org/downloads/"
  echo
  read -r -p "按回车键关闭窗口..." _
  exit 1
fi

printf '当前 Python：%s  (%s)\n\n' "$("$PY" -V 2>&1)" "$(command -v "$PY")"

if [ $# -eq 0 ]; then
  echo "稍后在提示处把文件夹直接拖进本窗口（或粘贴完整路径），再回车即可。"
  echo "结果会放在该文件夹下的 md_<当天日期>/ 里，转换完自动打开结果文件夹。"
  echo
fi

"$PY" ./files_to_md.py "$@" --open-result
code=$?

echo
echo "------------------------------------------------------------"
printf '转换结束（退出码 %s），原文件均未改动。\n' "$code"
read -r -p "按回车键关闭窗口..." _
exit "$code"
