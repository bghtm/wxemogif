# -*- coding: utf-8 -*-
"""wxemogif —— 图片批量转 GIF 小工具（重点：透明背景不丢）

功能
----
* 拖拽：把图片 / 文件夹直接拖进窗口
* 导入：按钮选择文件或文件夹（文件夹会递归扫描图片）
* 粘贴：Ctrl+V 粘贴剪切板里的图片（截图、浏览器复制的图）或复制的文件
* 批量转换：一次把列表里所有图片转成 GIF，输出到「源文件同目录的同名 .gif」

透明度说明
----------
GIF 只支持 1 位透明度（一个像素要么全透明、要么不透明），所以：
* 最多用 255 种颜色 + 1 个「透明色」索引，颜色质量最大化；
* 不把半透明像素和黑色/白色做混合（那样会出现黑边/白边光晕），
  而是保留像素本身的 RGB，只让 alpha 低于阈值的像素变成透明；
* 多帧动画统一使用同一个透明索引，并设置 disposal=2 + background=透明索引，
  避免播放时透明区域被背景色填充。

依赖
----
Python 3.8+ 与 Pillow；拖拽用 Win32 API 实现，不需要额外第三方库。

用法
----
    python wxemogif.py                # 打开窗口
    python wxemogif.py a.png b.jpg    # 打开窗口并预先加入这些文件
    python wxemogif.py --cli *.png    # 无界面批量转换（调试/脚本用）
"""

from __future__ import annotations

import argparse
import os
import queue
import sys
import threading
import time
import traceback
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    from PIL import Image, ImageDraw, ImageGrab, ImageStat
except ImportError:  # pragma: no cover - 缺依赖时给出友好提示
    try:
        sys.stderr.write("缺少 Pillow，请先执行：pip install Pillow\n")
    except Exception:
        pass
    raise

APP_NAME = "wxemogif"
APP_TITLE = "wxemogif · 图片批量转 GIF（保留透明）"
APP_VERSION = "1.1"

# 可选：OLE 拖放支持（收浏览器拖过来的网页图片）。缺失时只影响网页拖放。
try:
    import ole_dnd
except Exception:  # pragma: no cover
    ole_dnd = None

# Pillow 版本兼容：9.1 起这些常量被挪进了枚举，两头都兼容一下
try:
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:  # pragma: no cover - 老版本 Pillow
    RESAMPLE_LANCZOS = Image.LANCZOS
try:
    QUANTIZE_MEDIANCUT = Image.Quantize.MEDIANCUT
except AttributeError:  # pragma: no cover - 老版本 Pillow
    QUANTIZE_MEDIANCUT = Image.MEDIANCUT

# 扫描文件夹时认为可能是图片的扩展名
IMAGE_EXTS = {
    ".png", ".apng", ".jpg", ".jpeg", ".jpe", ".jfif", ".bmp", ".dib",
    ".webp", ".tif", ".tiff", ".tga", ".ico", ".ppm", ".pgm", ".pbm",
    ".pnm", ".avif", ".jp2", ".j2k", ".pcx", ".psd", ".gif",
}


# ---------------------------------------------------------------------------
# 1. Windows 拖拽支持（WM_DROPFILES + SetWindowSubclass）
# ---------------------------------------------------------------------------
class DropTarget:
    """把 tkinter 窗口（含所有子控件）注册成 Windows 拖放目标。

    Tk 在 Windows 上会为每个控件建一个真实 HWND（Label 甚至是原生 Static 控件），
    而拖放命中测试命中的是鼠标下面「最深层」的那个 HWND，所以必须递归地把整棵
    控件树都注册上，否则拖到列表、标签上面就不会触发。

    重要：窗口过程（WNDPROC）里**绝对不要调用任何 Tk 接口**（包括 root.after），
    也不要 import 东西。在 CPython 3.13 + _tkinter 的组合下，Tcl 派发窗口消息时
    会先释放 GIL，此时 ctypes 回调里再重入 Tcl 会把线程状态搞坏，之后事件循环里
    跑的任意 Python 回调都会以
        Fatal Python error: PyEval_RestoreThread: ... thread state is NULL
    崩溃。所以这里只做「纯数据」的事：把文件名丢进一个 list，剩下的交给主线程用
    root.after 轮询处理（见 App._tick）。
    """

    WM_DROPFILES = 0x0233
    SUBCLASS_ID = 0x5758  # 'WX'

    def __init__(self, root) -> None:
        import ctypes
        from ctypes import wintypes

        self._ctypes = ctypes
        self._wintypes = wintypes
        self.root = root
        self._pending: List[tuple] = []
        self.enabled = False
        self._registered: List[int] = []
        self._proc = None
        self.ole = None
        self._use_ole = False
        self._ole_log: List[str] = []
        self._user32 = ctypes.WinDLL("user32", use_last_error=True)
        self._shell32 = ctypes.WinDLL("shell32", use_last_error=True)
        self._comctl32 = ctypes.WinDLL("comctl32", use_last_error=True)

        LRESULT = ctypes.c_ssize_t
        WPARAM = ctypes.c_size_t
        LPARAM = ctypes.c_ssize_t
        UINT_PTR = ctypes.c_size_t
        DWORD_PTR = ctypes.c_size_t
        self._WPARAM, self._LPARAM = WPARAM, LPARAM

        self._SUBCLASSPROC = ctypes.WINFUNCTYPE(
            LRESULT, wintypes.HWND, ctypes.c_uint, WPARAM, LPARAM, UINT_PTR, DWORD_PTR
        )

        self._user32.PostMessageW.argtypes = [
            wintypes.HWND, ctypes.c_uint, WPARAM, LPARAM,
        ]
        self._comctl32.SetWindowSubclass.restype = wintypes.BOOL
        self._comctl32.SetWindowSubclass.argtypes = [
            wintypes.HWND, self._SUBCLASSPROC, UINT_PTR, DWORD_PTR,
        ]
        self._comctl32.DefSubclassProc.restype = LRESULT
        self._comctl32.DefSubclassProc.argtypes = [
            wintypes.HWND, ctypes.c_uint, WPARAM, LPARAM,
        ]
        self._shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        self._shell32.DragQueryFileW.argtypes = [
            wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT,
        ]
        self._shell32.DragQueryFileW.restype = wintypes.UINT
        self._shell32.DragFinish.argtypes = [wintypes.HANDLE]

    # -- 生命周期 ---------------------------------------------------------
    def enable(self) -> bool:
        """注册整棵控件树；应在本窗口所有控件创建完成后调用。

        优先用 OLE 的 IDropTarget：只有它能收到浏览器从网页拖过来的图片。
        如果 OLE 不可用，再退回 WM_DROPFILES（至少保证资源管理器拖文件能用）。
        """
        if self.enabled:
            return True

        if ole_dnd is not None:
            try:
                self.ole = ole_dnd.OleDropTarget(log=self._ole_log.append)
                if self.ole.initialize():
                    self._use_ole = True
                else:
                    self.ole = None
            except Exception:
                self.ole = None
                self._use_ole = False

        self._proc = self._SUBCLASSPROC(self._on_message)
        self._walk(self.root)
        self.enabled = bool(self._registered)
        return self.enabled

    def ole_diagnostics(self) -> str:
        if self._use_ole and self.ole is not None:
            return f"网页拖放：已启用（OLE 注册 {self.ole.registered_count} 个窗口）"
        return "网页拖放：不可用（已退回 WM_DROPFILES，浏览器拖图片可能无效）"

    def refresh(self) -> None:
        """新增控件后重新扫描（已注册的会跳过）。只能在主线程调用。"""
        if self._proc is not None:
            self._walk(self.root)

    def disable(self) -> None:
        if self.ole is not None:
            try:
                self.ole.revoke_all()
                self.ole.uninitialize()
            except Exception:
                pass
        if not self.enabled:
            return
        for hwnd in self._registered:
            try:
                self._comctl32.RemoveWindowSubclass(
                    hwnd, self._proc, self.SUBCLASS_ID
                )
            except Exception:
                pass
        self._registered.clear()
        self.enabled = False

    # -- 给主线程取数据 ---------------------------------------------------
    def take_pending(self) -> List[tuple]:
        """取走攒下的拖放数据（只在主线程调用）。

        返回 payload 列表，见 ole_dnd 的说明：
            ("files", [路径...]) ("blob", bytes, 名字)
            ("image", PIL.Image, 名字) ("url", 链接)
        """
        items: List[tuple] = []
        if self.ole is not None:
            try:
                items.extend(self.ole.take_pending())
            except Exception:
                pass
        if self._pending:
            paths, self._pending = self._pending, []
            items.append(("files", paths))
        return items

    # -- 内部：窗口过程里只碰纯数据和 Win32 ---------------------------------
    def _walk(self, widget) -> None:
        try:
            hwnd = widget.winfo_id()
        except Exception:
            return
        if hwnd and hwnd not in self._registered:
            self._registered.append(hwnd)
            # 优先把窗口注册成 OLE 拖放目标；成功就不再挂 WM_DROPFILES，
            # 因为 WS_EX_ACCEPTFILES 会让外壳改用 WM_DROPFILES，反而收不到网页拖放。
            registered_ole = False
            if self._use_ole and self.ole is not None:
                try:
                    registered_ole = self.ole.register(hwnd)
                except Exception:
                    registered_ole = False
            if not registered_ole:
                self._comctl32.SetWindowSubclass(hwnd, self._proc, self.SUBCLASS_ID, 0)
                self._shell32.DragAcceptFiles(hwnd, True)
        try:
            children = widget.winfo_children()
        except Exception:
            return
        for child in children:
            self._walk(child)

    def _query_files(self, hdrop) -> List[str]:
        shell32 = self._shell32
        count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        names: List[str] = []
        for i in range(count):
            length = shell32.DragQueryFileW(hdrop, i, None, 0)
            buf = self._ctypes.create_unicode_buffer(length + 1)
            shell32.DragQueryFileW(hdrop, i, buf, length + 1)
            names.append(buf.value)
        return names

    def _on_message(self, hwnd, msg, wparam, lparam, uid, ref):
        if msg == self.WM_DROPFILES:
            try:
                names = self._query_files(wparam)
                if names:
                    self._pending.extend(names)  # 纯 Python 操作，不碰 Tk
            except Exception:
                pass
            finally:
                try:
                    self._shell32.DragFinish(wparam)
                except Exception:
                    pass
            return 0
        return self._comctl32.DefSubclassProc(hwnd, msg, wparam, lparam)


# ---------------------------------------------------------------------------
# 2. 转换核心
# ---------------------------------------------------------------------------
class ConvertOptions:
    """转换参数。"""

    def __init__(
        self,
        keep_alpha: bool = True,
        alpha_threshold: int = 128,
        matte_color: Tuple[int, int, int] = (255, 255, 255),
        max_colors: int = 255,
        overwrite: bool = True,
    ) -> None:
        self.keep_alpha = bool(keep_alpha)
        self.alpha_threshold = max(1, min(255, int(alpha_threshold)))
        self.matte_color = tuple(int(c) for c in matte_color[:3])
        # GIF 一共 256 个索引：留 1 个给透明色，其余给真实颜色
        self.max_colors = max(2, min(255, int(max_colors)))
        self.overwrite = bool(overwrite)


class ConvertResult:
    __slots__ = ("ok", "skipped", "message", "out_path", "frames", "out_size")

    def __init__(
        self,
        ok: bool = False,
        skipped: bool = False,
        message: str = "",
        out_path: str = "",
        frames: int = 0,
        out_size: int = 0,
    ) -> None:
        self.ok = ok
        self.skipped = skipped
        self.message = message
        self.out_path = out_path
        self.frames = frames
        self.out_size = out_size


def _mean_opaque_color(rgb: Image.Image, opaque_mask: Image.Image) -> Tuple[int, int, int]:
    """取不透明区域的平均色，用来填充「即将变透明」的像素，避免污染调色板。"""
    try:
        if opaque_mask.getextrema()[1] == 0:
            return (255, 255, 255)
        mean = ImageStat.Stat(rgb, mask=opaque_mask).mean
        return tuple(min(255, max(0, int(round(c)))) for c in mean[:3])
    except Exception:
        return (255, 255, 255)


def frame_to_palette(
    rgba: Image.Image,
    opts: ConvertOptions,
    *,
    reserve_transparent: bool,
    file_has_alpha: bool,
) -> Tuple[Image.Image, Optional[int]]:
    """把一帧 RGBA 转成 P 模式，必要时保留一个透明索引。

    返回 (P 模式图像, 透明索引或 None)。
    """
    if rgba.mode != "RGBA":
        rgba = rgba.convert("RGBA")

    if not opts.keep_alpha:
        # 用户要求不留透明：把图像合成到指定背景色上
        bg = Image.new("RGBA", rgba.size, opts.matte_color + (255,))
        flat = Image.alpha_composite(bg, rgba).convert("RGB")
        return flat.quantize(colors=256, method=QUANTIZE_MEDIANCUT), None

    colors = opts.max_colors
    idx = colors  # 预留给透明色的索引

    if not file_has_alpha:
        # 整份文件都没有透明像素：256 色全用上
        return rgba.convert("RGB").quantize(
            colors=256, method=QUANTIZE_MEDIANCUT
        ), None

    alpha = rgba.getchannel("A")
    _low, high = alpha.getextrema()

    if high == 0:
        # 整帧全透明：全部像素都指向透明索引
        blank = Image.new("P", rgba.size, idx)
        blank.putpalette([0, 0, 0] * (idx + 1))
        return blank, (idx if reserve_transparent else None)

    # 低于阈值的像素判为透明，其余保留原色
    # （关键：不与黑/白背景做混合，所以不会出现黑边、白边光晕）
    transparent_mask = alpha.point(lambda v: 255 if v < opts.alpha_threshold else 0)

    rgb = rgba.convert("RGB")
    if transparent_mask.getextrema()[1] != 0:
        # 把要变透明的像素先涂成不透明区域的平均色，
        # 让有限的调色板名额都花在真正看得见的颜色上
        opaque_mask = transparent_mask.point(lambda v: 255 - v)
        fill = _mean_opaque_color(rgb, opaque_mask)
        quant_input = rgb.copy()
        quant_input.paste(fill, mask=transparent_mask)
    else:
        # 这一帧没有透明像素，但整份文件有：保持同样的索引布局
        quant_input = rgb

    palette_img = quant_input.quantize(colors=colors, method=QUANTIZE_MEDIANCUT)

    # 保证调色板长度正好是 (idx+1)*3，并把 idx 号位置留白给透明色
    palette = list(palette_img.getpalette() or [])
    palette = (palette + [0, 0, 0] * (idx + 1))[: idx * 3] + [0, 0, 0]
    palette_img.putpalette(palette)

    if transparent_mask.getextrema()[1] != 0:
        palette_img.paste(idx, mask=transparent_mask)
        return palette_img, (idx if reserve_transparent else None)

    return palette_img, None


def _load_frames(src: str) -> Tuple[List[Image.Image], List[int], int]:
    """读取源图的所有帧（RGBA）、每帧时长、循环次数。"""
    frames: List[Image.Image] = []
    durations: List[int] = []
    loop = 0
    with Image.open(src) as im:
        try:
            loop = int(im.info.get("loop", 0))
        except Exception:
            loop = 0
        try:
            frame_count = int(getattr(im, "n_frames", 1) or 1)
        except Exception:
            frame_count = 1
        for i in range(frame_count):
            try:
                im.seek(i)
            except EOFError:
                break
            except Exception:
                if i == 0:
                    raise
                break
            frames.append(im.convert("RGBA"))
            try:
                duration = int(im.info.get("duration", 100))
            except Exception:
                duration = 100
            durations.append(max(20, duration))
    if not frames:
        raise ValueError("没有读到任何图像帧")
    return frames, durations, loop


def convert_image(src: str, dst: str, opts: ConvertOptions) -> ConvertResult:
    """把 src 转成 GIF 写到 dst。"""
    if os.path.splitext(src)[1].lower() == ".gif":
        return ConvertResult(skipped=True, message="已经是 GIF，跳过", out_path=dst)

    if os.path.exists(dst) and not opts.overwrite:
        return ConvertResult(skipped=True, message="目标已存在，未覆盖", out_path=dst)

    if os.path.abspath(src) == os.path.abspath(dst):
        return ConvertResult(skipped=True, message="源文件与目标同名，跳过", out_path=dst)

    frames, durations, loop = _load_frames(src)

    # 整份文件只要有任意一帧存在透明像素，就让所有帧统一保留透明索引：
    # 这样动画每帧调色板布局一致，disposal / background 的语义也一致。
    file_has_alpha = False
    if opts.keep_alpha:
        for rgba in frames:
            if rgba.getchannel("A").getextrema()[0] < opts.alpha_threshold:
                file_has_alpha = True
                break

    pal_frames: List[Image.Image] = []
    transparents: List[Optional[int]] = []
    for rgba in frames:
        p_img, t_idx = frame_to_palette(
            rgba, opts,
            reserve_transparent=file_has_alpha,
            file_has_alpha=file_has_alpha,
        )
        pal_frames.append(p_img)
        transparents.append(t_idx)

    first = pal_frames[0]
    if transparents[0] is not None:
        # 逐帧写 info，让 Pillow 知道每帧自己的透明索引
        first.info["transparency"] = transparents[0]
    for p_img, t_idx in zip(pal_frames[1:], transparents[1:]):
        if t_idx is not None:
            p_img.info["transparency"] = t_idx

    save_kwargs = dict(
        format="GIF",
        save_all=True,
        append_images=pal_frames[1:],
        duration=durations if len(durations) > 1 else durations[0],
        loop=loop,
        optimize=False,   # 保住手工构造的调色板和透明索引，不让 Pillow 改写
        disposal=2,       # 每帧绘制前恢复背景，透明区域才不会被上一帧盖住
    )
    if file_has_alpha:
        # 逻辑屏幕背景索引 = 透明索引槽位（每帧都预留了同一个位置），
        # 免得解码器拿一个不透明颜色去填透明区域
        save_kwargs["background"] = (
            transparents[0] if transparents[0] is not None else opts.max_colors
        )

    tmp_path = dst + ".wxemogif.tmp"
    try:
        first.save(tmp_path, **save_kwargs)
        os.replace(tmp_path, dst)   # 先写临时文件再替换，中途失败不会毁掉已有文件
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise
    finally:
        for img in frames:
            try:
                img.close()
            except Exception:
                pass
        for img in pal_frames:
            try:
                img.close()
            except Exception:
                pass

    return ConvertResult(
        ok=True,
        message="完成",
        out_path=dst,
        frames=len(pal_frames),
        out_size=os.path.getsize(dst),
    )


def output_path_for(src: str, out_dir: Optional[str] = None,
                    stem: Optional[str] = None) -> str:
    """按「同目录同名 .gif」规则算出输出路径。

    stem 用来覆盖输出名：网页/剪切板来的图片落盘时会带一串时间戳后缀，
    但输出名应该还原成用户看到的原始文件名。
    """
    if stem is None:
        stem = os.path.splitext(os.path.basename(src))[0]
    folder = out_dir if out_dir else os.path.dirname(os.path.abspath(src))
    return os.path.join(folder, stem + ".gif")


def copy_as_is(src: str, dst: str, overwrite: bool = True) -> ConvertResult:
    """源文件本来就是 GIF 时直接复制过去。

    网页上拖下来的动图、剪切板里的 GIF 都走这条路：再编码一遍没有任何意义，
    但也不能像本地文件那样「跳过」，否则用户拖进来会什么都得不到。
    """
    import shutil

    if os.path.exists(dst) and not overwrite:
        return ConvertResult(skipped=True, message="目标已存在，未覆盖", out_path=dst)
    try:
        shutil.copy2(src, dst)
    except Exception as exc:
        return ConvertResult(message=f"复制失败：{exc}", out_path=dst)
    return ConvertResult(ok=True, message="已是 GIF，直接复制", out_path=dst,
                         frames=0, out_size=os.path.getsize(dst))


def is_image_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def collect_images(folder: str, recursive: bool = True) -> List[str]:
    """扫描文件夹里的图片文件。"""
    found: List[str] = []
    if recursive:
        for root_dir, dirnames, filenames in os.walk(folder):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for name in sorted(filenames):
                if is_image_file(name):
                    found.append(os.path.join(root_dir, name))
    else:
        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if os.path.isfile(full) and is_image_file(name):
                found.append(full)
    return found


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024.0
    return f"{num:.1f} GB"


def default_output_dir() -> str:
    """默认输出目录：桌面。

    走 Windows 的 Known Folder API，这样 OneDrive 把桌面重定向过也能拿到
    真正的桌面路径；失败再退回 %USERPROFILE%\\Desktop。
    """
    if sys.platform == "win32":
        try:
            import ctypes
            import uuid
            from ctypes import wintypes

            class GUID(ctypes.Structure):
                _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                            ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

            # FOLDERID_Desktop
            u = uuid.UUID("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")
            fid = GUID(u.fields[0], u.fields[1], u.fields[2],
                       (ctypes.c_ubyte * 8)(*u.bytes[8:]))
            shell32 = ctypes.WinDLL("shell32", use_last_error=True)
            ole32 = ctypes.WinDLL("ole32", use_last_error=True)
            shell32.SHGetKnownFolderPath.argtypes = [
                ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE,
                ctypes.POINTER(ctypes.c_wchar_p)]
            shell32.SHGetKnownFolderPath.restype = ctypes.c_long
            ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
            ptr = ctypes.c_wchar_p()
            hr = shell32.SHGetKnownFolderPath(ctypes.byref(fid), 0, None,
                                              ctypes.byref(ptr))
            if hr == 0 and ptr.value:
                path = ptr.value
                try:
                    ole32.CoTaskMemFree(ctypes.cast(ptr, ctypes.c_void_p))
                except Exception:
                    pass
                if os.path.isdir(path):
                    return path
        except Exception:
            pass
    home = os.path.expanduser("~")
    for name in ("Desktop", "桌面", "OneDrive/Desktop", "OneDrive/桌面"):
        path = os.path.join(home, *name.split("/"))
        if os.path.isdir(path):
            return path
    return home or os.getcwd()


# ---------------------------------------------------------------------------
# 3. 图形界面
# ---------------------------------------------------------------------------
def enable_dpi_awareness() -> float:
    """开启 DPI 感知并返回缩放比，避免高分屏下字体发虚。"""
    if sys.platform != "win32":
        return 1.0
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
        except Exception:
            dpi = 96
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


def _preload_pil() -> None:
    """在主线程提前把 PIL 的解码插件 import 完，避免工作线程里首次 import。"""
    try:
        Image.preinit()
        Image.init()
    except Exception:
        pass


class Task:
    """一次转换任务。

    没有列表，所以不需要 iid / 状态列之类的东西：拖进来一张就生成一个 Task，
    交给后台线程立刻转，转完把结果丢回界面显示。
    """

    __slots__ = ("path", "name", "dst", "opts", "is_temp", "status", "out_path", "note")

    def __init__(self, path: str, dst: str, opts: ConvertOptions,
                 is_temp: bool = False, display_name: Optional[str] = None) -> None:
        self.path = path
        self.name = display_name or os.path.basename(path)
        self.dst = dst                  # 输出路径：在主线程就定好，工作线程不再碰 Tk
        self.opts = opts                # 转换参数快照，避免工作线程读 tkinter 变量
        self.is_temp = is_temp
        self.status = "pending"
        self.out_path = ""
        self.note = ""


class App:
    def __init__(self, root, initial_files: Sequence[str] = ()) -> None:
        import tkinter as tk
        from tkinter import font as tkfont
        from tkinter import ttk

        self.tk = tk
        self.ttk = ttk
        self.root = root

        # 只保留「正在处理」的临时队列，界面上不显示，拖一张转一张
        self._queue: List[Task] = []
        self._lock = threading.Lock()
        self._worker_alive = False
        self._batch_total = 0
        self._batch_done = 0
        self.stats = {"ok": 0, "skip": 0, "fail": 0}

        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._last_source: Optional[str] = None
        self._last_output: Optional[str] = None
        self._last_result: Optional[ConvertResult] = None
        self._last_dir = ""
        self._paste_seq = 0
        self._src_photo = None
        self._out_photo = None
        self._preview_job = None
        self._preview_cache = {}

        root.title(APP_TITLE)
        root.geometry("980x760")
        root.minsize(860, 640)

        # -- 外观 ---------------------------------------------------------
        style = ttk.Style()
        try:
            style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        except Exception:
            pass
        default_font = tkfont.nametofont("TkDefaultFont")
        for family in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI"):
            try:
                default_font.configure(family=family, size=10)
                break
            except Exception:
                continue
        for named in ("TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(named).configure(
                    family=default_font.cget("family"), size=10)
            except Exception:
                pass

        # -- 变量 ---------------------------------------------------------
        self.keep_alpha = tk.BooleanVar(value=True)
        self.threshold = tk.IntVar(value=128)
        self.matte = (255, 255, 255)
        self.overwrite = tk.BooleanVar(value=True)
        self.same_dir = tk.BooleanVar(value=False)
        self.out_dir = tk.StringVar(value=default_output_dir())
        self.status_text = tk.StringVar(
            value="就绪：把图片拖进来就会立刻转换（也可以按 Ctrl+V 粘贴）")

        self._build_ui()
        self._bind_events()

        self.drop = None
        try:
            self.drop = DropTarget(root)
            self.drop.enable()
        except Exception:
            self.log("⚠ 拖拽功能初始化失败，请改用「添加图片」按钮或 Ctrl+V 粘贴")
            traceback.print_exc()
        if self.drop and not self.drop.enabled:
            self.log("⚠ 拖拽目标注册失败，请改用「添加图片」按钮或 Ctrl+V 粘贴")

        self.log("GIF 只支持 1 位透明度：alpha 小于阈值的像素会变全透明，"
                 "其余保留原色（所以不会产生黑边）")
        if self.drop is not None:
            self.log(self.drop.ole_diagnostics())
        self.update_preview()

        if initial_files:
            self.root.after(50, lambda: self.submit_paths(initial_files, source="启动参数"))

        # 常驻心跳：取走拖拽数据 + 刷新工作线程事件（原因见 DropTarget 注释）
        self._tick()

    # ------------------------------------------------------------------
    # 界面搭建
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        tk, ttk = self.tk, self.ttk

        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        # --- 工具栏 ---
        bar = ttk.Frame(outer)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(bar, text="添加图片…", command=self.choose_files, width=14).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", command=self.choose_folder, width=14).pack(
            side="left", padx=(6, 0))
        ttk.Button(bar, text="粘贴剪切板 (Ctrl+V)", command=self.paste_clipboard,
                   width=20).pack(side="left", padx=(6, 0))
        self.open_btn = ttk.Button(bar, text="打开输出文件夹", command=self.open_out_dir,
                                   width=16, state="disabled")
        self.open_btn.pack(side="right")
        self.count_label = ttk.Label(bar, text="已转换 0 张")
        self.count_label.pack(side="right", padx=(0, 12))

        # --- 预览：原图 vs 转换结果 ---
        preview = ttk.Labelframe(
            outer, text=" 预览（棋盘格 = 透明区域，拖入即转换） ", padding=8)
        preview.grid(row=1, column=0, sticky="nsew")
        preview.columnconfigure(0, weight=1, uniform="pv")
        preview.columnconfigure(1, weight=1, uniform="pv")
        preview.rowconfigure(1, weight=1)

        ttk.Label(preview, text="原图", anchor="center").grid(row=0, column=0, sticky="ew")
        ttk.Label(preview, text="转换结果（.gif）", anchor="center").grid(
            row=0, column=1, sticky="ew")

        self.src_canvas = tk.Canvas(preview, height=260, bg="#f4f4f4",
                                    highlightthickness=1, highlightbackground="#cccccc")
        self.src_canvas.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.out_canvas = tk.Canvas(preview, height=260, bg="#f4f4f4",
                                    highlightthickness=1, highlightbackground="#cccccc")
        self.out_canvas.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        for canvas in (self.src_canvas, self.out_canvas):
            canvas.bind("<Configure>", self._on_preview_resize)

        self.preview_info = ttk.Label(preview, text="", anchor="w", justify="left")
        self.preview_info.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        # --- 选项 ---
        opts = ttk.Labelframe(outer, text=" 选项 ", padding=8)
        opts.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        ttk.Checkbutton(opts, text="保留透明背景", variable=self.keep_alpha,
                        command=self._toggle_alpha).grid(row=0, column=0, sticky="w")
        ttk.Label(opts, text="透明阈值：").grid(row=0, column=1, sticky="e", padx=(16, 2))
        self.thr_spin = ttk.Spinbox(opts, from_=1, to=255, width=5,
                                    textvariable=self.threshold)
        self.thr_spin.grid(row=0, column=2, sticky="w")
        ttk.Label(opts, text="(alpha 小于它的像素变全透明，128 为常用值)").grid(
            row=0, column=3, sticky="w", padx=(6, 0))

        ttk.Label(opts, text="不保留透明时填充：").grid(
            row=1, column=1, sticky="e", padx=(16, 2), pady=(6, 0))
        self.matte_btn = tk.Button(opts, text="  ", bg="#ffffff", width=4,
                                   relief="groove", command=self.choose_matte)
        self.matte_btn.grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Checkbutton(opts, text="覆盖已存在的 .gif", variable=self.overwrite).grid(
            row=1, column=3, sticky="w", padx=(6, 0), pady=(6, 0))

        # --- 输出位置 ---
        outbox = ttk.Labelframe(outer, text=" 输出位置 ", padding=8)
        outbox.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        outbox.columnconfigure(2, weight=1)
        ttk.Radiobutton(outbox, text="输出到指定文件夹（默认：桌面）：", value=False,
                        variable=self.same_dir, command=self._toggle_outdir).grid(
            row=0, column=0, sticky="w")
        self.out_entry = ttk.Entry(outbox, textvariable=self.out_dir)
        self.out_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(4, 4))
        self.out_btn = ttk.Button(outbox, text="浏览…", command=self.choose_outdir, width=8)
        self.out_btn.grid(row=0, column=3, sticky="e")
        ttk.Radiobutton(outbox, text="输出到图片所在文件夹（同名 .gif）", value=True,
                        variable=self.same_dir, command=self._toggle_outdir).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(outbox, text="提示：从网页/剪切板拖进来或粘进来的图片没有源目录，"
                               "一律输出到上面这个文件夹。",
                  foreground="#777777").grid(row=2, column=0, columnspan=4, sticky="w",
                                             pady=(6, 0))
        self._toggle_outdir()

        # --- 状态 ---
        status_row = ttk.Frame(outer)
        status_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        status_row.columnconfigure(0, weight=1)
        ttk.Label(status_row, textvariable=self.status_text, anchor="w").grid(
            row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(status_row, mode="determinate", maximum=1, length=160)
        self.progress.grid(row=0, column=1, sticky="e", padx=(8, 0))

        # --- 日志 ---
        logframe = ttk.Labelframe(outer, text=" 日志 ", padding=4)
        logframe.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.log_text = tk.Text(logframe, height=8, wrap="word", state="disabled",
                                background="#fbfbfb", relief="flat")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb = ttk.Scrollbar(logframe, orient="vertical", command=self.log_text.yview)
        log_sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_sb.set)

    def _bind_events(self) -> None:
        self.root.bind("<Control-v>", lambda e: self.paste_clipboard())
        self.root.bind("<Control-V>", lambda e: self.paste_clipboard())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.report_callback_exception = self._report_exception

    def _report_exception(self, exc_type, exc_value, exc_tb) -> None:
        text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        self.log("✖ 界面异常：\n" + text.strip())

    # ------------------------------------------------------------------
    # 控件回调
    # ------------------------------------------------------------------
    def _toggle_alpha(self) -> None:
        self.thr_spin.configure(state="normal" if self.keep_alpha.get() else "disabled")

    def _toggle_outdir(self) -> None:
        state = "disabled" if self.same_dir.get() else "normal"
        self.out_entry.configure(state=state)
        self.out_btn.configure(state=state)

    def choose_matte(self) -> None:
        from tkinter import colorchooser

        rgb, _ = colorchooser.askcolor(color="#%02x%02x%02x" % self.matte,
                                       title="选择填充背景色")
        if rgb:
            self.matte = tuple(int(c) for c in rgb)
            self.matte_btn.configure(bg="#%02x%02x%02x" % self.matte)

    def choose_files(self) -> None:
        from tkinter import filedialog

        patterns = " ".join(f"*{e}" for e in sorted(IMAGE_EXTS))
        paths = filedialog.askopenfilenames(
            title="选择图片（可多选，选完立刻转换）",
            filetypes=[("图片文件", patterns), ("所有文件", "*.*")],
        )
        if paths:
            self.submit_paths(list(paths), source="选择文件")

    def choose_folder(self) -> None:
        from tkinter import filedialog

        folder = filedialog.askdirectory(title="选择包含图片的文件夹（会全部转换）")
        if folder:
            self.submit_paths([folder], source="选择文件夹")

    def choose_outdir(self) -> None:
        from tkinter import filedialog

        folder = filedialog.askdirectory(title="选择输出文件夹",
                                         initialdir=self.out_dir.get() or os.getcwd())
        if folder:
            self.out_dir.set(folder)

    def open_out_dir(self) -> None:
        folder = self._last_dir or (self.out_dir.get() if not self.same_dir.get() else "")
        if folder and os.path.isdir(folder):
            try:
                os.startfile(folder)  # noqa: S606 - Windows 专用
            except Exception as exc:
                self.log(f"✖ 无法打开文件夹：{exc}")
        else:
            self.log("⚠ 还没有输出文件夹可打开")

    # ------------------------------------------------------------------
    # 剪切板
    # ------------------------------------------------------------------
    def paste_clipboard(self) -> None:
        try:
            data = ImageGrab.grabclipboard()
        except Exception as exc:
            self.log(f"✖ 读取剪切板失败：{exc}")
            self.status_text.set("读取剪切板失败")
            return

        if isinstance(data, list):
            paths = [p for p in data if isinstance(p, str)]
            if paths:
                self.submit_paths(paths, source="剪切板")
                return

        if isinstance(data, Image.Image):
            self._add_clipboard_image(data)
            return

        # 有些程序复制的是「文件路径文本」
        try:
            text = self.root.clipboard_get()
        except Exception:
            text = ""
        if text:
            candidate = text.strip().strip('"')
            if os.path.isfile(candidate):
                self.submit_paths([candidate], source="剪切板")
                return

        self.log("⚠ 剪切板里没有图片（用截图工具截图后直接 Ctrl+V 即可）")
        self.status_text.set("剪切板里没有图片")

    def _add_clipboard_image(self, image: Image.Image) -> None:
        """剪切板里的位图没有名字，给个固定的 clipboard 名字，输出名才好看。"""
        self._paste_seq += 1
        name = "clipboard.png" if self._paste_seq == 1 else f"clipboard_{self._paste_seq}.png"
        self._handle_payloads([("image", image, name)])

    # ------------------------------------------------------------------
    # 拖放数据 -> 任务
    # ------------------------------------------------------------------
    def _temp_dir(self) -> str:
        import tempfile

        folder = os.path.join(tempfile.gettempdir(), "wxemogif_paste")
        try:
            os.makedirs(folder, exist_ok=True)
            return folder
        except OSError:
            return tempfile.gettempdir()

    @staticmethod
    def _safe_stem(name: str) -> str:
        stem = os.path.splitext(os.path.basename(name or ""))[0]
        cleaned = "".join(ch if (ch.isalnum() or ch in "-_") else "_" for ch in stem)
        return (cleaned.strip("_") or "image")[:40]

    def _save_temp_bytes(self, data: bytes, filename: str) -> Optional[str]:
        """把拖进来/下载来的原始字节落成临时文件（保留原扩展名，动画不会丢）。"""
        ext = os.path.splitext(filename or "")[1].lower()
        if ext not in IMAGE_EXTS:
            ext = ".png"
        path = os.path.join(
            self._temp_dir(),
            f"{self._safe_stem(filename)}_{time.time_ns() % 10 ** 8}{ext}")
        try:
            with open(path, "wb") as fh:
                fh.write(data)
        except OSError as exc:
            self.log(f"✖ 写入临时文件失败：{exc}")
            return None
        return path

    def _save_temp_image(self, image: Image.Image, filename: str) -> Optional[str]:
        path = os.path.join(
            self._temp_dir(),
            f"{self._safe_stem(filename)}_{time.time_ns() % 10 ** 8}.png")
        try:
            image.convert("RGBA").save(path, "PNG")
        except Exception as exc:
            self.log(f"✖ 保存拖放的图片失败：{exc}")
            return None
        return path

    def _handle_payloads(self, payloads: Iterable[tuple]) -> None:
        """处理一批拖放数据（见 DropTarget.take_pending）。"""
        real_paths: List[str] = []
        temp_paths: List[str] = []
        source_names: dict = {}
        used_names: set = set()

        def unique_name(name: str) -> str:
            """同一批里重名就加序号，免得两个网页图片同名互相覆盖。"""
            stem, ext = os.path.splitext(name)
            candidate, index = name, 1
            while candidate.lower() in used_names:
                candidate = f"{stem}_{index}{ext}"
                index += 1
            used_names.add(candidate.lower())
            return candidate

        for item in payloads:
            if not item:
                continue
            kind = item[0]
            if kind == "files":
                real_paths.extend(item[1])
            elif kind == "blob":
                _, data, name = item[:3]
                path = self._save_temp_bytes(data, unique_name(name))
                if path:
                    temp_paths.append(path)
                    source_names[path] = name
                    self.log(f"＋ 网页图片「{name}」已接收（{human_size(len(data))}）")
            elif kind == "image":
                _, img, name = item[:3]
                path = self._save_temp_image(img, unique_name(name))
                if path:
                    temp_paths.append(path)
                    source_names[path] = name
                    self.log(f"＋ 拖放的位图已接收（{img.width}×{img.height}）")
            elif kind == "url":
                self.log(f"… 拖放只给了图片地址，正在下载：{item[1]}")
                self._fetch_url_async(item[1])

        if real_paths:
            self.submit_paths(real_paths, source="拖入")
        if temp_paths:
            self.submit_paths(temp_paths, source="拖入", is_temp=True,
                              source_names=source_names)

    def _fetch_url_async(self, url: str) -> None:
        threading.Thread(target=self._fetch_url, args=(url,), daemon=True).start()

    def _fetch_url(self, url: str) -> None:
        """兜底路径：拖放数据里只有 URL 时，把图片下载下来。"""
        try:
            import urllib.request

            request = urllib.request.Request(
                url, headers={"User-Agent": "Mozilla/5.0 (wxemogif)"})
            with urllib.request.urlopen(request, timeout=20) as response:
                data = response.read(128 * 1024 * 1024)
        except Exception as exc:
            self._events.put(("log", f"✖ 下载图片失败：{exc}"))
            return
        if not data:
            return
        name = os.path.basename(url.split("?")[0].split("#")[0]) or "web_image.png"
        path = self._save_temp_bytes(data, name)
        if path:
            self._events.put(("add_temp", path, f"网页图片 {name}"))

    # ------------------------------------------------------------------
    # 任务提交与执行
    # ------------------------------------------------------------------
    def submit_paths(self, paths: Iterable[str], source: str = "",
                     is_temp: bool = False, source_names: Optional[dict] = None) -> None:
        """把一批文件直接变成任务并立刻开始转换（不攒列表）。"""
        if not self.same_dir.get():
            out_dir = self.out_dir.get().strip()
            if not out_dir or not os.path.isdir(out_dir):
                self.log(f"✖ 输出文件夹不存在：{out_dir}")
                self.status_text.set("输出文件夹不存在，请重新选择")
                return

        resolved: List[str] = []
        for raw in paths:
            if not raw:
                continue
            path = os.path.abspath(raw)
            if os.path.isdir(path):
                found = collect_images(path, recursive=True)
                self.log(f"＋ 文件夹 {os.path.basename(path)}：找到 {len(found)} 张图片")
                resolved.extend(found)
            elif os.path.isfile(path):
                resolved.append(path)
            else:
                self.log(f"⚠ 忽略不存在的路径：{path}")

        tasks: List[Task] = []
        for path in resolved:
            if not is_image_file(path):
                self.log(f"⚠ 跳过非图片文件：{os.path.basename(path)}")
                continue
            name = (source_names or {}).get(path)
            tasks.append(self._make_task(path, is_temp, name))

        if not tasks:
            self.status_text.set("没有可转换的图片")
            return

        self.log(f"＋ {source or '导入'}：{len(tasks)} 张，立刻转换")
        self._enqueue(tasks)

    def _make_task(self, path: str, is_temp: bool = False,
                   display_name: Optional[str] = None) -> Task:
        """在主线程把输出路径和参数定下来，工作线程就只用干纯活。"""
        name = display_name or os.path.basename(path)
        stem = os.path.splitext(name)[0] or "image"
        if is_temp or not self.same_dir.get():
            out_dir = self.out_dir.get().strip() or os.path.dirname(path)
        else:
            out_dir = os.path.dirname(path)
        return Task(path, output_path_for(path, out_dir, stem=stem), self._options(),
                    is_temp=is_temp, display_name=name)

    def _enqueue(self, tasks: List[Task]) -> None:
        if not tasks:
            return
        with self._lock:
            self._queue.extend(tasks)
            self._batch_total += len(tasks)
            start = not self._worker_alive
            if start:
                self._worker_alive = True

        # 预览先切到这一张的原图，用户马上能看到转的是哪张
        self._last_source = tasks[0].path
        self._last_output = None
        self._last_result = None
        self.update_preview()

        if start:
            threading.Thread(target=self._worker_loop, daemon=True).start()
        else:
            self.log(f"… 上一张还在转，这 {len(tasks)} 张排在后面")
        self.progress.configure(maximum=max(1, self._batch_total), value=self._batch_done)
        self.status_text.set(f"正在转换…（{self._batch_done}/{self._batch_total}）")

    def _worker_loop(self) -> None:
        while True:
            with self._lock:
                if not self._queue:
                    self._worker_alive = False
                    self._events.put(("idle",))
                    return
                task = self._queue.pop(0)
            try:
                if task.is_temp and task.path.lower().endswith(".gif"):
                    # 网页 / 剪切板来的本来就是 GIF，直接复制过去
                    result = copy_as_is(task.path, task.dst, task.opts.overwrite)
                else:
                    result = convert_image(task.path, task.dst, task.opts)
            except Exception as exc:
                result = ConvertResult(message=f"{type(exc).__name__}: {exc}")
            self._events.put(("result", task, result))

    # ------------------------------------------------------------------
    # 预览
    # ------------------------------------------------------------------
    def _checkerboard(self, size: Tuple[int, int], cell: int = 8):
        key = (size, cell)
        cached = self._preview_cache.get(key)
        if cached is not None:
            return cached
        width, height = size
        board = Image.new("RGB", (width, height), (255, 255, 255))
        draw = ImageDraw.Draw(board)
        for y in range(0, height, cell):
            for x in range(0, width, cell):
                if ((x // cell) + (y // cell)) % 2:
                    draw.rectangle((x, y, x + cell - 1, y + cell - 1), fill=(205, 205, 205))
        if len(self._preview_cache) > 8:
            self._preview_cache.clear()
        self._preview_cache[key] = board
        return board

    def _on_preview_resize(self, _event=None) -> None:
        """预览面板尺寸变化后重绘（节流，免得拖窗口时反复缩放图片）。"""
        if self._preview_job is not None:
            try:
                self.root.after_cancel(self._preview_job)
            except Exception:
                pass
        self._preview_job = self.root.after(150, self.update_preview)

    def _render_canvas(self, canvas, path: Optional[str], empty_text: str):
        """把图片（棋盘格衬底）画到 canvas 上，返回 PhotoImage（调用方要保住引用）。"""
        from PIL import ImageTk

        width = max(80, canvas.winfo_width())
        height = max(80, canvas.winfo_height())
        canvas.delete("all")
        if not path or not os.path.exists(path):
            if empty_text:
                canvas.create_text(width // 2, height // 2, text=empty_text,
                                   fill="#9a9a9a", justify="center")
            return None
        try:
            with Image.open(path) as im:
                if getattr(im, "n_frames", 1) > 1:
                    im.seek(0)
                im.thumbnail((width, height), RESAMPLE_LANCZOS)
                rgba = im.convert("RGBA")
        except Exception as exc:
            canvas.create_text(width // 2, height // 2, text=f"无法预览：{exc}",
                               fill="#c0392b", width=width - 16)
            return None
        board = self._checkerboard(rgba.size).copy()
        board.paste(rgba, (0, 0), rgba)
        photo = ImageTk.PhotoImage(board)
        canvas.create_image(width // 2, height // 2, image=photo, anchor="center")
        return photo

    def update_preview(self) -> None:
        self._src_photo = self._render_canvas(
            self.src_canvas, self._last_source, "把图片拖进来\n（或 Ctrl+V 粘贴）")

        show_output = (self._last_result is not None and self._last_result.ok
                       and self._last_output and os.path.exists(self._last_output))
        self._out_photo = self._render_canvas(
            self.out_canvas, self._last_output if show_output else None,
            "转换结果会显示在这里")

        self.preview_info.configure(text=self._describe())

    def _describe(self) -> str:
        parts: List[str] = []
        if self._last_source and os.path.exists(self._last_source):
            try:
                with Image.open(self._last_source) as im:
                    frames = int(getattr(im, "n_frames", 1) or 1)
                    alpha = ""
                    if im.mode in ("RGBA", "LA", "P", "PA"):
                        if im.convert("RGBA").getchannel("A").getextrema()[0] < 255:
                            alpha = "，有透明"
                    text = (f"原图 {os.path.basename(self._last_source)}："
                            f"{im.size[0]}×{im.size[1]}")
                    if frames > 1:
                        text += f"，{frames} 帧"
                    parts.append(text + alpha)
            except Exception:
                pass
        result = self._last_result
        if result is not None:
            if result.ok:
                text = (f"输出 {os.path.basename(result.out_path)}："
                        f"{human_size(result.out_size)}")
                if result.frames > 1:
                    text += f"，{result.frames} 帧"
                if result.message and result.message != "完成":
                    text += f"（{result.message}）"
                parts.append(text)
            elif result.message:
                parts.append(result.message)
        return "\n".join(parts) or "把图片拖进来就会立刻转成 GIF，输出到上面设定的文件夹"

    # ------------------------------------------------------------------
    # 日志
    # ------------------------------------------------------------------
    def log(self, message: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{stamp}] {message}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except Exception:
            try:
                print(f"[{stamp}] {message}")
            except Exception:
                pass  # --windowed 打包后没有控制台

    # ------------------------------------------------------------------
    # 心跳 / 事件
    # ------------------------------------------------------------------
    def _tick(self) -> None:
        """常驻心跳：取走拖拽数据 + 刷新工作线程事件。只在主线程跑。"""
        try:
            if self.drop is not None:
                pending = self.drop.take_pending()
                if pending:
                    self._handle_payloads(pending)
        except Exception:
            traceback.print_exc()
        try:
            self._drain_events()
        except Exception:
            traceback.print_exc()
        try:
            self.root.after(100, self._tick)
        except Exception:
            pass  # 窗口已销毁

    def _drain_events(self) -> None:
        while True:
            try:
                event = self._events.get_nowait()
            except queue.Empty:
                return
            kind = event[0]
            if kind == "result":
                _, task, result = event
                self._on_result(task, result)
            elif kind == "idle":
                self._on_idle()
            elif kind == "add_temp":
                _, path, label = event
                self.submit_paths([path], source=label or "网页图片", is_temp=True)
            elif kind == "log":
                self.log(event[1])

    def _on_result(self, task: Task, result: ConvertResult) -> None:
        self._batch_done += 1
        self.progress.configure(maximum=max(1, self._batch_total), value=self._batch_done)
        self._last_source = task.path
        self._last_result = result

        if result.ok:
            self.stats["ok"] += 1
            task.status = "done"
            task.out_path = result.out_path
            self._last_output = result.out_path
            self._last_dir = os.path.dirname(result.out_path)
            self.open_btn.configure(state="normal")
            extra = f"（{result.frames} 帧）" if result.frames > 1 else ""
            self.log(f"✔ {task.name} → {result.out_path} {extra}· {human_size(result.out_size)}")
            self.status_text.set(
                f"完成：{os.path.basename(result.out_path)} → {os.path.dirname(result.out_path)}")
        elif result.skipped:
            self.stats["skip"] += 1
            task.status = "skip"
            self._last_output = None
            self.log(f"－ {task.name}：{result.message}")
            self.status_text.set(f"跳过 {task.name}：{result.message}")
        else:
            self.stats["fail"] += 1
            task.status = "fail"
            self._last_output = None
            self.log(f"✖ {task.name} 转换失败：{result.message}")
            self.status_text.set(f"失败：{result.message}")

        done = self.stats["ok"]
        extra = []
        if self.stats["skip"]:
            extra.append(f"跳过 {self.stats['skip']}")
        if self.stats["fail"]:
            extra.append(f"失败 {self.stats['fail']}")
        self.count_label.configure(
            text=f"已转换 {done} 张" + (f"（{'，'.join(extra)}）" if extra else ""))
        self.update_preview()

    def _on_idle(self) -> None:
        self._batch_total = 0
        self._batch_done = 0
        self.progress.configure(maximum=1, value=0)
        if self.stats["ok"] or self.stats["skip"] or self.stats["fail"]:
            self.status_text.set(
                f"就绪：拖入图片立刻转换（累计完成 {self.stats['ok']} 张）")
        else:
            self.status_text.set("就绪：把图片拖进来就会立刻转换")

    # ------------------------------------------------------------------
    # 转换参数
    # ------------------------------------------------------------------
    def _options(self) -> ConvertOptions:
        return ConvertOptions(
            keep_alpha=self.keep_alpha.get(),
            alpha_threshold=self.threshold.get(),
            matte_color=self.matte,
            overwrite=self.overwrite.get(),
        )

    def on_close(self) -> None:
        with self._lock:
            busy = self._worker_alive
        if busy:
            from tkinter import messagebox

            if not messagebox.askokcancel("正在转换", "还有图片在转换中，确定要退出吗？"):
                return
        if self.drop is not None:
            self.drop.disable()
        self.root.destroy()


# ---------------------------------------------------------------------------
# 4. 命令行模式 / 入口
# ---------------------------------------------------------------------------
def _ensure_console() -> None:
    """--windowed 打包出来的 exe 没有控制台；用 --cli 时挂到调用它的父控制台上。"""
    if sys.platform != "win32" or sys.stdout is not None:
        return
    try:
        import ctypes

        ATTACH_PARENT_PROCESS = -1
        if not ctypes.windll.kernel32.AttachConsole(ATTACH_PARENT_PROCESS):
            return
        sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
        sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
    except Exception:
        pass


def run_cli(args: argparse.Namespace) -> int:
    # 打包成无控制台的 exe 时，先把自己挂回父窗口的控制台，否则什么都打印不出来
    _ensure_console()
    # Windows 控制台默认是 GBK，遇到打不出的字符不要让程序崩掉
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    opts = ConvertOptions(
        keep_alpha=not args.no_alpha,
        alpha_threshold=args.threshold,
        matte_color=(255, 255, 255),
        overwrite=not args.no_overwrite,
    )
    targets: List[str] = []
    for raw in args.cli_files:
        if os.path.isdir(raw):
            targets.extend(collect_images(raw, recursive=True))
        else:
            targets.append(os.path.abspath(raw))
    if not targets:
        print("没有找到要转换的图片")
        return 1
    ok = skipped = failed = 0
    for src in targets:
        dst = output_path_for(src, args.outdir)
        try:
            result = convert_image(src, dst, opts)
        except Exception as exc:
            failed += 1
            print(f"[失败] {os.path.basename(src)}: {type(exc).__name__}: {exc}")
            continue
        if result.ok:
            ok += 1
            print(f"[完成] {os.path.basename(src)} -> {result.out_path} "
                  f"({result.frames} 帧, {human_size(result.out_size)})")
        else:
            skipped += 1
            print(f"[跳过] {os.path.basename(src)}: {result.message}")
    print(f"\n完成 {ok}，跳过 {skipped}，失败 {failed}")
    return 0 if failed == 0 else 2


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="wxemogif",
        description="图片批量转 GIF（保留透明背景），图形界面 / 命令行两用",
    )
    parser.add_argument("files", nargs="*", help="启动时预先加入列表的图片或文件夹")
    parser.add_argument("--cli", action="store_true",
                        help="不打开界面，直接转换（配合 files 使用）")
    parser.add_argument("--outdir", default=None,
                        help="输出文件夹（--cli 模式默认与源文件同目录；"
                             "带界面运行时用它覆盖默认的桌面）")
    parser.add_argument("--threshold", type=int, default=128,
                        help="透明阈值 1-255，默认 128")
    parser.add_argument("--no-alpha", action="store_true",
                        help="不保留透明，改为填充背景色")
    parser.add_argument("--no-overwrite", action="store_true",
                        help="不覆盖已存在的 .gif")
    args = parser.parse_args(list(argv) if argv is not None else None)
    args.cli_files = args.files

    if args.cli:
        return run_cli(args)

    _preload_pil()
    scale = enable_dpi_awareness()
    import tkinter as tk

    root = tk.Tk()
    try:
        if scale > 1.0:
            root.tk.call("tk", "scaling", scale * 96.0 / 72.0)
    except Exception:
        pass
    app = App(root, initial_files=args.files)
    if args.outdir:
        # 带界面运行时也允许预设输出目录（做快捷方式很方便）
        folder = os.path.abspath(args.outdir)
        if os.path.isdir(folder):
            app.out_dir.set(folder)
            app.same_dir.set(False)
            app._toggle_outdir()
        else:
            app.log(f"⚠ --outdir 指向的文件夹不存在，继续用默认输出位置：{folder}")
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
