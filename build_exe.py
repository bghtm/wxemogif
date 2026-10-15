# -*- coding: utf-8 -*-
"""把 wxemogif 打包成单文件 exe。

用法：
    python build_exe.py            # 打包（默认单文件 + 无控制台窗口）
    python build_exe.py --onedir   # 打包成一个文件夹（启动更快）
    python build_exe.py --console  # 保留控制台窗口（调试用）

产出：dist/wxemogif.exe
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(ROOT, "wxemogif.py")
ICON = os.path.join(ROOT, "app.ico")

# 确保这些解码器被收进 exe（Pillow 的 hook 一般能自动带上，这里再显式兜底）
HIDDEN_IMPORTS = [
    "ole_dnd",              # 网页拖放支持，和主程序同目录
    "PIL.WebPImagePlugin",
    "PIL.PngImagePlugin",
    "PIL.JpegImagePlugin",
    "PIL.GifImagePlugin",
    "PIL.BmpImagePlugin",
    "PIL.IcoImagePlugin",
    "PIL.TiffImagePlugin",
    "PIL.PpmImagePlugin",
    "PIL.PsdImagePlugin",
    "PIL.TgaImagePlugin",
    "PIL.PcxImagePlugin",
    "PIL.ImageGrab",
]


def make_icon(path: str = ICON) -> bool:
    """随手生成一个带透明背景的表情图标（多尺寸 ico）。"""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return False
    try:
        ss = 8
        size = 256 * ss
        big = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(big)
        m = 10 * ss

        # 圆脸
        d.ellipse((m, m, size - m, size - m), fill=(255, 205, 60, 255))
        # 眼睛
        d.ellipse((72 * ss, 84 * ss, 106 * ss, 124 * ss), fill=(58, 40, 22, 255))
        d.ellipse((150 * ss, 84 * ss, 184 * ss, 124 * ss), fill=(58, 40, 22, 255))
        # 笑脸
        d.arc((78 * ss, 108 * ss, 178 * ss, 186 * ss), 15, 165,
              fill=(58, 40, 22, 255), width=11 * ss)
        # GIF 的“透明”暗示：左上角一块棋盘格切角
        cell = 16 * ss
        for row in range(6):
            for col in range(6):
                if (row + col) % 2 == 0:
                    x0, y0 = col * cell, row * cell
                    d.rectangle((x0, y0, x0 + cell - 1, y0 + cell - 1),
                                fill=(255, 255, 255, 210))
        # 切掉左上角，露出透明
        d.polygon([(0, 0), (96 * ss, 0), (0, 96 * ss)], fill=(0, 0, 0, 0))

        icon = big.resize((256, 256), Image.LANCZOS)
        icon.save(path, format="ICO",
                  sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
        print(f"[icon] 已生成 {path}")
        return True
    except Exception as exc:
        print(f"[icon] 生成失败（继续打包，只是没有图标）：{exc}")
        return False


def main() -> int:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("没有安装 PyInstaller，正在安装…")
        if subprocess.call([sys.executable, "-m", "pip", "install", "pyinstaller"]) != 0:
            print("PyInstaller 安装失败")
            return 1

    onedir = "--onedir" in sys.argv
    console = "--console" in sys.argv

    if not make_icon():
        pass

    for folder in ("build", "dist"):
        full = os.path.join(ROOT, folder)
        if os.path.isdir(full):
            shutil.rmtree(full, ignore_errors=True)

    # PyInstaller 自己会在 build/ 里放工作文件，用独立 spec 名避免污染
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--name", "wxemogif",
        "--onedir" if onedir else "--onefile",
        "--console" if console else "--windowed",
        "--specpath", ROOT,
        "--workpath", os.path.join(ROOT, "build"),
        "--distpath", os.path.join(ROOT, "dist"),
    ]
    if os.path.exists(ICON):
        cmd += ["--icon", ICON]
    for name in HIDDEN_IMPORTS:
        cmd += ["--hidden-import", name]
    # 界面里的中文提示都是代码内的，不需要额外数据文件；把源码目录排除掉
    cmd += ["--exclude-module", "tkinter.test", "--exclude-module", "test"]
    cmd += [APP]

    print("[build] " + " ".join(cmd))
    code = subprocess.call(cmd, cwd=ROOT)
    if code != 0:
        print("打包失败")
        return code

    exe = (os.path.join(ROOT, "dist", "wxemogif", "wxemogif.exe") if onedir
           else os.path.join(ROOT, "dist", "wxemogif.exe"))
    if os.path.exists(exe):
        print(f"\n完成：{exe}  ({os.path.getsize(exe) / 1048576:.1f} MB)")
    else:
        print("\n打包结束，但没有找到预期的 exe，请看上面的输出")
    return 0


if __name__ == "__main__":
    sys.exit(main())
