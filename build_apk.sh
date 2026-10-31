#!/usr/bin/env bash
# ============================================================
#  一键编译 APK（Linux / WSL / macOS）
#  用法：
#     ./build_apk.sh            # 编 arm64 + arm32 两个包（兼容老手机，慢一些）
#     ./build_apk.sh arm64      # 只编 arm64（快一半，现在绝大多数手机都是 arm64）
# ============================================================
set -e

cd "$(dirname "$0")"

ARCHS="arm64-v8a,armeabi-v7a"
if [ "$1" = "arm64" ]; then
    ARCHS="arm64-v8a"
    echo ">> 只编译 arm64（体积更小、速度更快）"
fi

echo "=============================================="
echo "  第 1/4 步：安装系统依赖（需要 sudo 密码）"
echo "=============================================="
if command -v apt-get >/dev/null 2>&1; then
    sudo apt-get update
    sudo apt-get install -y --no-install-recommends \
        git zip unzip autoconf libtool pkg-config \
        zlib1g-dev libncurses-dev cmake libffi-dev libssl-dev \
        make
    # JDK 17 优先；老版本 Ubuntu 没有 17 时依次降级
    sudo apt-get install -y openjdk-17-jdk \
        || sudo apt-get install -y openjdk-11-jdk \
        || sudo apt-get install -y default-jdk \
        || { echo "!! JDK 安装失败，请手动安装 Java 后重跑"; exit 1; }
else
    echo "!! 没找到 apt-get。请手动安装 git / zip / unzip / JDK17 后重跑本脚本。"
    exit 1
fi

echo
echo "=============================================="
echo "  第 2/4 步：安装 buildozer"
echo "=============================================="
python3 -m pip install --upgrade pip
python3 -m pip install --upgrade buildozer "cython==3.0.11"

# macOS 上 pip 装到用户目录时要把路径加进 PATH
export PATH="$HOME/.local/bin:$PATH"

echo
echo "=============================================="
echo "  第 3/4 步：设定 CPU 架构 -> $ARCHS"
echo "=============================================="
sed -i.bak "s/^android.archs = .*/android.archs = $ARCHS/" buildozer.spec

echo
echo "=============================================="
echo "  第 4/4 步：编译（首次会下载 SDK/NDK 约 1.5GB，"
echo "             大约 20~40 分钟，请耐心等待）"
echo "=============================================="
buildozer -v android debug

echo
echo "=============================================="
echo "  完成！APK 在 bin/ 目录："
ls -lh bin/*.apk 2>/dev/null || echo "  （没找到 apk，请往上翻看红色报错）"
echo "=============================================="
