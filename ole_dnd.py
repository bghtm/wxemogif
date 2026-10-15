# -*- coding: utf-8 -*-
"""OLE 拖放支持（Windows）：接收浏览器从网页上拖过来的图片。

为什么需要这个
--------------
Explorer 里拖文件走的是 `WM_DROPFILES`，但**浏览器拖网页上的图片走的是 OLE 的
`IDropTarget`**。所以只处理 `WM_DROPFILES` 的话，从网页往窗口里拖图片毫无反应。
这里用 ctypes 手写一个 `IDropTarget`，并把数据对象里浏览器可能提供的格式都读出来：

======================================  ============================================
格式                                     说明
======================================  ============================================
`CF_HDROP`                               真实文件路径（Explorer 拖文件）
`FileGroupDescriptorW` + `FileContents`  虚拟文件——浏览器拖图片就是这种，
                                         能拿到**原始字节**（PNG 的透明通道、
                                          动图的每一帧都不会丢）
`CF_DIB` / `CF_BITMAP`                   位图（退而求其次）
`CF_UNICODETEXT` / `CF_HTML`            图片 URL（最后的兜底，交给上层去下载）
======================================  ============================================

线程与重入
----------
`DragEnter` / `Drop` 这些回调是被 OLE 从 Tcl 事件循环里调进来的，所以和窗口过程
一样**只做纯数据操作**：`Drop` 里只 `AddRef` 一下数据对象指针就立刻返回，真正的
解析留给主线程轮询（`take_pending`）。这样既不会卡住拖放循环，也避开了
「ctypes 回调里重入 Tcl 会搞坏线程状态」的坑。

对外接口
--------
    t = OleDropTarget()
    t.initialize()        # 初始化 OLE（STA），失败返回 False
    t.register(hwnd)      # 把窗口注册成拖放目标，失败返回 False
    t.take_pending()      # 主线程调用，返回 payload 列表
    t.revoke_all()        # 撤销注册
    t.uninitialize()

payload 形式：
    ("files", [路径, ...])           真实文件
    ("blob",  bytes, 文件名)          虚拟文件的原始字节（网页图片）
    ("image", PIL.Image, 文件名)      位图
    ("url",   "https://...")         图片 URL
"""

from __future__ import annotations

import ctypes
import io
import os
import struct
import sys
import uuid
from ctypes import wintypes
from typing import List, Optional, Tuple

if sys.platform != "win32":  # pragma: no cover - 本工具只在 Windows 上用
    raise ImportError("ole_dnd 只支持 Windows")

# ---------------------------------------------------------------------------
# 基础类型
# ---------------------------------------------------------------------------
HRESULT = ctypes.c_long
ULONG = ctypes.c_ulong
DWORD = ctypes.c_ulong
LONG = ctypes.c_long
WORD = ctypes.c_ushort
LPSTR = ctypes.c_char_p

DVASPECT_CONTENT = 1
TYMED_HGLOBAL = 1
TYMED_ISTREAM = 4
TYMED_GDI = 16

DROPEFFECT_NONE = 0
DROPEFFECT_COPY = 1
DROPEFFECT_MOVE = 2
DROPEFFECT_LINK = 4

S_OK = 0
E_NOINTERFACE = -2147467262  # 0x80004002
E_NOTIMPL = -2147467263      # 0x80004001

CF_TEXT = 1
CF_BITMAP = 2
CF_DIB = 8
CF_UNICODETEXT = 13
CF_HDROP = 15

IID_IUNKNOWN_LE = uuid.UUID("{00000000-0000-0000-C000-000000000046}").bytes_le
IID_IDROPTARGET_LE = uuid.UUID("{00000122-0000-0000-C000-000000000046}").bytes_le

# IDataObject 各方法在虚表里的下标（IUnknown 占 0/1/2）
IDATAOBJECT_ADDREF = 1
IDATAOBJECT_RELEASE = 2
IDATAOBJECT_GETDATA = 3
IDATAOBJECT_QUERYGETDATA = 5

# IStream 各方法下标
ISTREAM_READ = 4


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


class SIZEL(ctypes.Structure):
    _fields_ = [("cx", LONG), ("cy", LONG)]


class POINTL(ctypes.Structure):
    _fields_ = [("x", LONG), ("y", LONG)]


class FORMATETC(ctypes.Structure):
    _fields_ = [
        ("cfFormat", WORD),
        ("ptd", ctypes.c_void_p),
        ("dwAspect", DWORD),
        ("lindex", LONG),
        ("tymed", DWORD),
    ]


class _STGMEDIUM_UNION(ctypes.Union):
    _fields_ = [
        ("hGlobal", ctypes.c_void_p),
        ("pstm", ctypes.c_void_p),
        ("pstg", ctypes.c_void_p),
        ("hBitmap", ctypes.c_void_p),
        ("hMetaFilePict", ctypes.c_void_p),
        ("hEnhMetaFile", ctypes.c_void_p),
        ("hFileName", ctypes.c_void_p),
        ("pStream", ctypes.c_void_p),
        ("pStorage", ctypes.c_void_p),
        ("pFile", ctypes.c_void_p),
    ]


class STGMEDIUM(ctypes.Structure):
    _fields_ = [
        ("tymed", DWORD),
        ("u", _STGMEDIUM_UNION),
        ("pUnkForRelease", ctypes.c_void_p),
    ]


class FILEDESCRIPTORW(ctypes.Structure):
    _fields_ = [
        ("dwFlags", DWORD),
        ("clsid", GUID),
        ("sizel", SIZEL),
        ("pointl", POINTL),
        ("dwFileAttributes", DWORD),
        ("ftCreationTime", wintypes.FILETIME),
        ("ftLastAccessTime", wintypes.FILETIME),
        ("ftLastWriteTime", wintypes.FILETIME),
        ("nFileSizeHigh", DWORD),
        ("nFileSizeLow", DWORD),
        ("dwReserved", DWORD),
        ("cFileName", wintypes.WCHAR * 260),
    ]


class BITMAP(ctypes.Structure):
    _fields_ = [
        ("bmType", LONG), ("bmWidth", LONG), ("bmHeight", LONG),
        ("bmWidthBytes", LONG), ("bmBitsPixel", WORD), ("bmPlanes", WORD),
        ("bmBits", ctypes.c_void_p),
    ]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", DWORD), ("biWidth", LONG), ("biHeight", LONG),
        ("biPlanes", WORD), ("biBitCount", WORD), ("biCompression", DWORD),
        ("biSizeImage", DWORD), ("biXPelsPerMeter", LONG),
        ("biYPelsPerMeter", LONG), ("biClrUsed", DWORD), ("biClrImportant", DWORD),
    ]


# ---------------------------------------------------------------------------
# 动态库
# ---------------------------------------------------------------------------
_ole32 = ctypes.WinDLL("ole32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_shell32 = ctypes.WinDLL("shell32", use_last_error=True)
_user32 = ctypes.WinDLL("user32", use_last_error=True)
_gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

_ole32.OleInitialize.argtypes = [ctypes.c_void_p]
_ole32.OleInitialize.restype = HRESULT
_ole32.OleUninitialize.argtypes = []
_ole32.RegisterDragDrop.argtypes = [wintypes.HWND, ctypes.c_void_p]
_ole32.RegisterDragDrop.restype = HRESULT
_ole32.RevokeDragDrop.argtypes = [wintypes.HWND]
_ole32.RevokeDragDrop.restype = HRESULT
_ole32.ReleaseStgMedium.argtypes = [ctypes.POINTER(STGMEDIUM)]
_user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
_user32.RegisterClipboardFormatW.restype = wintypes.UINT

_kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
_kernel32.GlobalSize.restype = ctypes.c_size_t
_kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
_kernel32.GlobalLock.restype = ctypes.c_void_p
_kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]

_shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT,
                                    wintypes.LPWSTR, wintypes.UINT]
_shell32.DragQueryFileW.restype = wintypes.UINT

_gdi32.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT,
                             wintypes.UINT, ctypes.c_void_p, ctypes.c_void_p,
                             wintypes.UINT]
_gdi32.GetDIBits.restype = ctypes.c_int
_gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
_gdi32.CreateCompatibleDC.restype = wintypes.HDC
_gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
_gdi32.SelectObject.restype = wintypes.HGDIOBJ
_gdi32.DeleteDC.argtypes = [wintypes.HDC]
_gdi32.DeleteDC.restype = wintypes.BOOL
_gdi32.GetObjectW.argtypes = [wintypes.HGDIOBJ, ctypes.c_int, ctypes.c_void_p]
_gdi32.GetObjectW.restype = ctypes.c_int
_user32.GetDC.argtypes = [wintypes.HWND]
_user32.GetDC.restype = wintypes.HDC
_user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
_user32.ReleaseDC.restype = ctypes.c_int


# ---------------------------------------------------------------------------
# COM 辅助
# ---------------------------------------------------------------------------
def _vtbl_method(punk, index: int, restype, *argtypes):
    """从 COM 接口指针上取出第 index 个方法，包装成可调用对象。"""
    vtbl = ctypes.cast(punk, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    proto = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)
    return proto(vtbl[index])


def _com_addref(punk) -> int:
    fn = _vtbl_method(punk, IDATAOBJECT_ADDREF, ULONG)
    return fn(punk)


def _com_release(punk) -> int:
    fn = _vtbl_method(punk, IDATAOBJECT_RELEASE, ULONG)
    return fn(punk)


def _hglobal_bytes(hglobal) -> bytes:
    size = _kernel32.GlobalSize(hglobal)
    if not size:
        return b""
    ptr = _kernel32.GlobalLock(hglobal)
    if not ptr:
        return b""
    try:
        return ctypes.string_at(ptr, size)
    finally:
        _kernel32.GlobalUnlock(hglobal)


class _DropTargetVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("DragEnter", ctypes.c_void_p),
        ("DragOver", ctypes.c_void_p),
        ("DragLeave", ctypes.c_void_p),
        ("Drop", ctypes.c_void_p),
    ]


class _DropTargetObj(ctypes.Structure):
    _fields_ = [("lpVtbl", ctypes.POINTER(_DropTargetVtbl))]


# ---------------------------------------------------------------------------
# 主体
# ---------------------------------------------------------------------------
class OleDropTarget:
    """一个可以注册到多个窗口上的 IDropTarget。"""

    def __init__(self, log=None) -> None:
        self._log = log
        self._initialized = False
        self._refs = 1
        self._raw_pending: List[int] = []   # 攒下的 IDataObject 指针（Drop 里只做这个）
        self._registered: List[int] = []
        self._keepalive: List[object] = []

        # 浏览器可能用到的自定义剪贴板格式号（要在有窗口之前注册好）
        self._fmt_filegroupdescriptor = _user32.RegisterClipboardFormatW(
            "FileGroupDescriptorW")
        self._fmt_filecontents = _user32.RegisterClipboardFormatW("FileContents")
        self._fmt_html = _user32.RegisterClipboardFormatW("HTML Format")
        self._fmt_mozurl = _user32.RegisterClipboardFormatW("text/x-moz-url")

        self._build_vtable()

    # -- 生命周期 ---------------------------------------------------------
    def initialize(self) -> bool:
        """在当前线程初始化 OLE（STA）。必须在主线程、创建窗口之后调用。"""
        hr = _ole32.OleInitialize(None)
        # S_OK=0；S_FALSE=1 表示已经初始化过；其余都算失败
        if hr not in (S_OK, 1):
            self._log_msg(f"OLE 初始化失败 hr=0x{hr & 0xFFFFFFFF:08X}")
            return False
        self._initialized = True
        return True

    def uninitialize(self) -> None:
        if self._initialized:
            try:
                _ole32.OleUninitialize()
            except Exception:
                pass
            self._initialized = False

    def register(self, hwnd: int) -> bool:
        if not self._initialized or not hwnd:
            return False
        try:
            hr = _ole32.RegisterDragDrop(wintypes.HWND(hwnd), self._ptr)
        except Exception:
            return False
        if hr == S_OK:
            self._registered.append(hwnd)
            return True
        return False

    def revoke_all(self) -> None:
        for hwnd in self._registered:
            try:
                _ole32.RevokeDragDrop(wintypes.HWND(hwnd))
            except Exception:
                pass
        self._registered.clear()

    @property
    def registered_count(self) -> int:
        return len(self._registered)

    # -- 虚表 -------------------------------------------------------------
    def _build_vtable(self) -> None:
        proto_qi = ctypes.WINFUNCTYPE(HRESULT, ctypes.c_void_p, ctypes.c_void_p,
                                      ctypes.c_void_p)
        proto_ref = ctypes.WINFUNCTYPE(ULONG, ctypes.c_void_p)
        proto_enter = ctypes.WINFUNCTYPE(HRESULT, ctypes.c_void_p, ctypes.c_void_p,
                                         DWORD, POINTL, ctypes.POINTER(DWORD))
        proto_over = ctypes.WINFUNCTYPE(HRESULT, ctypes.c_void_p, DWORD, POINTL,
                                        ctypes.POINTER(DWORD))
        proto_leave = ctypes.WINFUNCTYPE(HRESULT, ctypes.c_void_p)
        proto_drop = ctypes.WINFUNCTYPE(HRESULT, ctypes.c_void_p, ctypes.c_void_p,
                                        DWORD, POINTL, ctypes.POINTER(DWORD))

        def qi(this, riid, ppv):
            try:
                if riid and ppv:
                    raw = ctypes.string_at(riid, 16)
                    if raw in (IID_IUNKNOWN_LE, IID_IDROPTARGET_LE):
                        self._refs += 1
                        ctypes.cast(ppv, ctypes.POINTER(ctypes.c_void_p))[0] = \
                            self._ptr.value
                        return S_OK
                    ctypes.cast(ppv, ctypes.POINTER(ctypes.c_void_p))[0] = None
            except Exception:
                pass
            return E_NOINTERFACE

        def addref(this):
            self._refs += 1
            return self._refs

        def release(this):
            self._refs -= 1
            return self._refs

        def drag_enter(this, pdata, keystate, pt, pdw_effect):
            effect = DROPEFFECT_NONE
            try:
                if pdata and self._has_usable_format(pdata):
                    effect = DROPEFFECT_COPY
            except Exception:
                effect = DROPEFFECT_COPY  # 拿不准就允许放下，别让用户白拖
            if pdw_effect:
                pdw_effect[0] = effect
            return S_OK

        def drag_over(this, keystate, pt, pdw_effect):
            if pdw_effect:
                pdw_effect[0] = DROPEFFECT_COPY
            return S_OK

        def drag_leave(this):
            return S_OK

        def drop(this, pdata, keystate, pt, pdw_effect):
            # 只 AddRef 存起来，真正的解析放到主线程，绝不在这里碰 Tk
            try:
                if pdata:
                    _com_addref(pdata)
                    self._raw_pending.append(pdata)
            except Exception:
                pass
            if pdw_effect:
                pdw_effect[0] = DROPEFFECT_COPY
            return S_OK

        self._fn_qi = proto_qi(qi)
        self._fn_addref = proto_ref(addref)
        self._fn_release = proto_ref(release)
        self._fn_enter = proto_enter(drag_enter)
        self._fn_over = proto_over(drag_over)
        self._fn_leave = proto_leave(drag_leave)
        self._fn_drop = proto_drop(drop)

        self._vtbl = _DropTargetVtbl()
        self._vtbl.QueryInterface = ctypes.cast(self._fn_qi, ctypes.c_void_p)
        self._vtbl.AddRef = ctypes.cast(self._fn_addref, ctypes.c_void_p)
        self._vtbl.Release = ctypes.cast(self._fn_release, ctypes.c_void_p)
        self._vtbl.DragEnter = ctypes.cast(self._fn_enter, ctypes.c_void_p)
        self._vtbl.DragOver = ctypes.cast(self._fn_over, ctypes.c_void_p)
        self._vtbl.DragLeave = ctypes.cast(self._fn_leave, ctypes.c_void_p)
        self._vtbl.Drop = ctypes.cast(self._fn_drop, ctypes.c_void_p)

        self._obj = _DropTargetObj()
        self._obj.lpVtbl = ctypes.pointer(self._vtbl)
        self._ptr = ctypes.cast(ctypes.byref(self._obj), ctypes.c_void_p)
        # 防止被 GC
        self._keepalive = [self._vtbl, self._obj, self._fn_qi, self._fn_addref,
                           self._fn_release, self._fn_enter, self._fn_over,
                           self._fn_leave, self._fn_drop]

    # -- 主线程侧 ---------------------------------------------------------
    def take_pending(self) -> List[tuple]:
        if not self._raw_pending:
            return []
        raws, self._raw_pending = self._raw_pending, []
        payloads: List[tuple] = []
        for pdata in raws:
            try:
                payloads.extend(self._extract(pdata))
            except Exception as exc:
                self._log_msg(f"解析拖放数据失败：{exc}")
            finally:
                try:
                    _com_release(pdata)
                except Exception:
                    pass
        return payloads

    # -- 数据对象读取 ------------------------------------------------------
    def _log_msg(self, text: str) -> None:
        if self._log:
            try:
                self._log(text)
            except Exception:
                pass

    def _get_data(self, pdata, cf: int, tymed: int, lindex: int = -1):
        """取一份数据，成功返回 STGMEDIUM，失败返回 None（调用方负责 Release）。"""
        fmt = FORMATETC(cf, None, DVASPECT_CONTENT, lindex, tymed)
        stg = STGMEDIUM()
        get_data = _vtbl_method(pdata, IDATAOBJECT_GETDATA, HRESULT,
                                ctypes.POINTER(FORMATETC),
                                ctypes.POINTER(STGMEDIUM))
        hr = get_data(pdata, ctypes.byref(fmt), ctypes.byref(stg))
        if hr != S_OK:
            return None
        return stg

    def _has_format(self, pdata, cf: int, tymed: int, lindex: int = -1) -> bool:
        fmt = FORMATETC(cf, None, DVASPECT_CONTENT, lindex, tymed)
        query = _vtbl_method(pdata, IDATAOBJECT_QUERYGETDATA, HRESULT,
                             ctypes.POINTER(FORMATETC))
        try:
            return query(pdata, ctypes.byref(fmt)) == S_OK
        except Exception:
            return False

    def _has_usable_format(self, pdata) -> bool:
        checks = (
            (CF_HDROP, TYMED_HGLOBAL),
            (self._fmt_filegroupdescriptor, TYMED_HGLOBAL),
            (CF_DIB, TYMED_HGLOBAL),
            (CF_DIB, TYMED_GDI),
            (CF_BITMAP, TYMED_GDI),
            (CF_UNICODETEXT, TYMED_HGLOBAL),
            (self._fmt_html, TYMED_HGLOBAL),
            (self._fmt_mozurl, TYMED_HGLOBAL),
            (CF_TEXT, TYMED_HGLOBAL),
        )
        for cf, tymed in checks:
            if cf and self._has_format(pdata, cf, tymed):
                return True
        return False

    def _release(self, stg) -> None:
        try:
            _ole32.ReleaseStgMedium(ctypes.byref(stg))
        except Exception:
            pass

    def _extract(self, pdata) -> List[tuple]:
        out: List[tuple] = []
        out.extend(self._extract_hdrop(pdata))
        if not out:
            out.extend(self._extract_virtual_files(pdata))
        if not out:
            out.extend(self._extract_bitmap(pdata))
        if not out:
            out.extend(self._extract_url(pdata))
        return out

    # -- 1) 真实文件（Explorer） -------------------------------------------
    def _extract_hdrop(self, pdata) -> List[tuple]:
        stg = self._get_data(pdata, CF_HDROP, TYMED_HGLOBAL)
        if stg is None:
            return []
        try:
            hdrop = stg.u.hGlobal
            count = _shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
            paths: List[str] = []
            for i in range(count):
                length = _shell32.DragQueryFileW(hdrop, i, None, 0)
                buf = ctypes.create_unicode_buffer(length + 1)
                _shell32.DragQueryFileW(hdrop, i, buf, length + 1)
                if buf.value:
                    paths.append(buf.value)
            return [("files", paths)] if paths else []
        finally:
            self._release(stg)

    # -- 2) 虚拟文件（浏览器拖图片的主要方式） -------------------------------
    def _extract_virtual_files(self, pdata) -> List[tuple]:
        stg = self._get_data(pdata, self._fmt_filegroupdescriptor, TYMED_HGLOBAL)
        if stg is None:
            return []
        try:
            blob = _hglobal_bytes(stg.u.hGlobal)
        finally:
            self._release(stg)
        if len(blob) < 4:
            return []

        count = struct.unpack_from("<I", blob, 0)[0]
        desc_size = ctypes.sizeof(FILEDESCRIPTORW)
        if count <= 0 or count > 64:
            return []

        out: List[tuple] = []
        for i in range(count):
            off = 4 + i * desc_size
            if off + desc_size > len(blob):
                break
            try:
                fd = FILEDESCRIPTORW.from_buffer_copy(blob[off:off + desc_size])
            except Exception:
                continue
            name = self._safe_name(fd.cFileName)
            data = self._read_file_contents(pdata, i)
            if not data:
                continue
            out.append(("blob", data, name))
        return out

    def _read_file_contents(self, pdata, index: int) -> bytes:
        # 多数实现（含 Chromium）用的是 IStream
        for tymed in (TYMED_ISTREAM, TYMED_HGLOBAL):
            stg = self._get_data(pdata, self._fmt_filecontents, tymed, lindex=index)
            if stg is None:
                continue
            try:
                if tymed == TYMED_ISTREAM:
                    return self._read_stream(stg.u.pstm)
                return _hglobal_bytes(stg.u.hGlobal)
            finally:
                self._release(stg)
        return b""

    def _read_stream(self, pstm) -> bytes:
        if not pstm:
            return b""
        read_fn = _vtbl_method(pstm, ISTREAM_READ, HRESULT, ctypes.c_void_p,
                               ULONG, ctypes.POINTER(ULONG))
        chunks: List[bytes] = []
        buf = ctypes.create_string_buffer(65536)
        got = ULONG(0)
        total = 0
        limit = 256 * 1024 * 1024  # 256 MB 上限，防御性
        while total < limit:
            hr = read_fn(pstm, buf, 65536, ctypes.byref(got))
            if hr != S_OK or got.value == 0:
                break
            chunks.append(buf.raw[:got.value])
            total += got.value
        return b"".join(chunks)

    # -- 3) 位图 -----------------------------------------------------------
    def _extract_bitmap(self, pdata) -> List[tuple]:
        from PIL import Image

        # 先试 CF_DIB（HGLOBAL，就是去掉文件头的 BMP）
        stg = self._get_data(pdata, CF_DIB, TYMED_HGLOBAL)
        if stg is not None:
            try:
                raw = _hglobal_bytes(stg.u.hGlobal)
            finally:
                self._release(stg)
            img = self._dib_to_image(raw)
            if img is not None:
                return [("image", img, "拖放的图片.png")]

        # 再试 GDI 位图句柄
        for cf in (CF_DIB, CF_BITMAP):
            stg = self._get_data(pdata, cf, TYMED_GDI)
            if stg is None:
                continue
            try:
                img = self._hbitmap_to_image(stg.u.hBitmap)
            finally:
                self._release(stg)
            if img is not None:
                return [("image", img, "拖放的图片.png")]
        return []

    def _dib_to_image(self, dib: bytes):
        """CF_DIB 的内容就是 BMP 去掉 14 字节文件头，把文件头补回去交给 Pillow。

        两个坑：
        * bfOffBits（像素数据起点）要算对：14(文件头) + biSize(信息头) + 调色板
          + 位域掩码。少算了 Pillow 会把信息头当像素读，颜色整个错位。
        * 32 位 BI_RGB 的位图，Pillow 会把第 4 个字节当填充丢掉，**透明度就没了**。
          而浏览器拖图片给出来的 CF_DIB 恰恰常是这种，所以这种格式自己解。
        """
        from PIL import Image

        if len(dib) < 16:
            return None
        try:
            bi_size = struct.unpack_from("<I", dib, 0)[0]
            width = struct.unpack_from("<i", dib, 4)[0]
            height = struct.unpack_from("<i", dib, 8)[0]
            bit_count = struct.unpack_from("<H", dib, 14)[0]
            compression = struct.unpack_from("<I", dib, 16)[0] if len(dib) >= 20 else 0

            if (bi_size >= 40 and bit_count == 32 and compression == 0
                    and 0 < width <= 20000 and height and abs(height) <= 20000):
                manual = self._decode_32bpp_dib(dib, bi_size, width, height)
                if manual is not None:
                    return manual

            if bi_size == 12:            # BITMAPCOREHEADER，调色板项是 3 字节
                entry_size = 3
                clr_used = 0
            else:
                entry_size = 4
                clr_used = struct.unpack_from("<I", dib, 32)[0] if len(dib) >= 36 else 0

            palette = 0
            if bit_count <= 8:
                palette = (clr_used or (1 << bit_count)) * entry_size
            elif clr_used:
                palette = clr_used * entry_size

            masks = 12 if (compression == 3 and bi_size == 40) else 0
            off_bits = 14 + bi_size + palette + masks

            header = struct.pack("<2sIHHI", b"BM", 14 + len(dib), 0, 0, off_bits)
            with Image.open(io.BytesIO(header + dib)) as im:
                im.load()
                return im.convert("RGBA")
        except Exception:
            return None

    @staticmethod
    def _decode_32bpp_dib(dib: bytes, bi_size: int, width: int, height: int):
        """自己解 32 位 BI_RGB 位图，把第 4 个字节当 alpha（GDI 那种全 0 的除外）。"""
        from PIL import Image

        bottom_up = height > 0
        rows_count = abs(height)
        stride = width * 4
        start = bi_size
        end = start + stride * rows_count
        if end > len(dib):
            return None
        try:
            blob = dib[start:end]
            if bottom_up:
                blob = b"".join(
                    blob[y * stride:(y + 1) * stride] for y in reversed(range(rows_count))
                )
            img = Image.frombytes("RGBA", (width, rows_count), blob,
                                  "raw", "BGRA", stride, 1)
            # alpha 整张一个值（常见是 GDI 位图的全 0）说明根本没有 alpha 通道
            low, high = img.getchannel("A").getextrema()
            if low == high:
                img = img.convert("RGB").convert("RGBA")
            return img
        except Exception:
            return None

    def _hbitmap_to_image(self, hbitmap):
        from PIL import Image

        if not hbitmap:
            return None
        hdc = None
        memdc = None
        try:
            hdc = _user32.GetDC(None)
            memdc = _gdi32.CreateCompatibleDC(hdc)
            _gdi32.SelectObject(memdc, hbitmap)

            bmp = BITMAP()
            if not _gdi32.GetObjectW(hbitmap, ctypes.sizeof(bmp), ctypes.byref(bmp)):
                return None
            width, height = int(bmp.bmWidth), int(bmp.bmHeight)
            if width <= 0 or height <= 0 or width > 20000 or height > 20000:
                return None

            bmi = BITMAPINFOHEADER()
            bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.biWidth = width
            bmi.biHeight = -height       # 负数 = top-down
            bmi.biPlanes = 1
            bmi.biBitCount = 32
            bmi.biCompression = 0        # BI_RGB

            stride = width * 4
            buf = ctypes.create_string_buffer(stride * height)
            got = _gdi32.GetDIBits(memdc, hbitmap, 0, height, buf,
                                   ctypes.byref(bmi), 0)
            if not got:
                return None
            img = Image.frombuffer("RGBA", (width, height), buf.raw,
                                   "raw", "BGRA", stride, 1).convert("RGBA")
            # GDI 位图常常没有 alpha（整张全 0），那样整张图会变全透明，
            # 所以 alpha 整张一个值时按不透明处理。
            low, high = img.getchannel("A").getextrema()
            if low == high:
                img = img.convert("RGB").convert("RGBA")
            return img
        except Exception:
            return None
        finally:
            try:
                if memdc:
                    _gdi32.DeleteDC(memdc)
            except Exception:
                pass
            try:
                if hdc:
                    _user32.ReleaseDC(None, hdc)
            except Exception:
                pass

    # -- 4) URL 兜底 --------------------------------------------------------
    def _extract_url(self, pdata) -> List[tuple]:
        # 每种格式的编码不一样：CF_HTML 是 UTF-8，其余 Unicode 格式是 UTF-16LE
        candidates = (
            (self._fmt_mozurl, "utf-16-le"),
            (CF_UNICODETEXT, "utf-16-le"),
            (self._fmt_html, "utf-8"),
            (CF_TEXT, "mbcs"),
        )
        for cf, encoding in candidates:
            if not cf:
                continue
            stg = self._get_data(pdata, cf, TYMED_HGLOBAL)
            if stg is None:
                continue
            try:
                raw = _hglobal_bytes(stg.u.hGlobal)
            finally:
                self._release(stg)
            if not raw:
                continue
            text = self._decode_text(raw, encoding)
            url = self._pick_image_url(text)
            if url:
                return [("url", url)]
        return []

    @staticmethod
    def _decode_text(raw: bytes, encoding: str) -> str:
        try:
            text = raw.decode(encoding, errors="ignore")
        except (LookupError, UnicodeDecodeError):
            text = raw.decode("utf-8", errors="ignore")
        # 去掉结尾的 NUL 以及后面的垃圾
        return text.split("\x00")[0]

    @staticmethod
    def _pick_image_url(text: str) -> Optional[str]:
        import re

        if not text:
            return None
        # text/x-moz-url 是「URL\n标题」
        for line in text.replace("\r\n", "\n").split("\n"):
            line = line.strip()
            if line.lower().startswith(("http://", "https://")) and \
                    re.search(r"\.(png|jpe?g|gif|webp|bmp|ico|avif)(\?|#|$)", line, re.I):
                return line
        # 从 HTML 片段里找 <img src=...>
        for match in re.finditer(r"""<img[^>]+src\s*=\s*["']([^"']+)["']""",
                                 text, re.I):
            src = match.group(1).strip()
            if src.lower().startswith(("http://", "https://")):
                return src
        # 整段就是一个图片地址
        stripped = text.strip().strip('"')
        if stripped.lower().startswith(("http://", "https://")) and \
                re.search(r"\.(png|jpe?g|gif|webp|bmp|ico|avif)(\?|#|$)", stripped, re.I):
            return stripped
        return None

    @staticmethod
    def _safe_name(name: str) -> str:
        name = (name or "").replace("\\", "/").split("/")[-1].strip()
        for ch in '<>:"|?*\x00':
            name = name.replace(ch, "_")
        name = name.lstrip(".") or "drag_image"
        return name[:120]
