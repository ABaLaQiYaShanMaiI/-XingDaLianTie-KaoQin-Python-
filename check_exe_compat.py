# -*- coding: utf-8 -*-
"""check_exe_compat.py — 打包产物兼容性自检(发布前务必跑一遍)

要回答的问题: 这个 exe 拷到别人的电脑上能不能启动?

做法: 读 exe 自身 + PyInstaller 包内所有 dll/pyd 的 PE 导入表, 找出引用了
      「Windows 8+ 才有的 API 集」(api-ms-win-core-* / api-ms-win-* 等) 的二进制。
      这类名字在 Windows 7 上既不是系统 DLL、也不在系统的 API 集表里, 加载器会直接
      弹「无法启动此程序, 因为计算机中丢失 xxx.dll」, 程序一行代码都跑不到
      (本项目的实盘事故: python313.dll / onnxruntime.dll 都要 api-ms-win-core-path-l1-1-0.dll)。

约定:
  * 只有 ``api-ms-win-crt-*`` 属于 UCRT, Windows 7 装上「VC++ 2015-2022 运行库」
    (KB2999226) 就能用 —— 不算致命;
  * 其余 ``api-ms-win-*`` / ``ext-ms-*`` 都是 Win8/Win10 才有的 API 集, Windows 7
    上直接起不来(除非单独打 KB2533623 一类补丁)。

用法:
    python check_exe_compat.py                     # 默认查 dist 下的 exe
    python check_exe_compat.py <exe 或 onedir 目录> [...]
    python check_exe_compat.py --all <exe>         # 连包内每个二进制的依赖都列出来
退出码: 0 = 可在 Windows 7 启动; 1 = 需要 Windows 8.1+; 2 = 分析失败。
"""
import os
import re
import struct
import sys
import tempfile

# Windows 7 上确实缺的 API 集(除 UCRT 外); UCRT 前缀单独放行
UCRT_PREFIX = 'api-ms-win-crt-'
APISET_RE = re.compile(r'^(?:api-ms-win|ext-ms-win)-', re.I)
# 打包用的 Python 版本 -> (官方最低系统, 说明)
PY_MIN_OS = {
    6: ('Windows 7 SP1', ''),
    7: ('Windows 7 SP1', ''),
    8: ('Windows 7 SP1', '需要 UCRT(VC++ 2015-2022 运行库 / KB2999226)'),
    9: ('Windows 8.1', '官方已不支持 Windows 7'),
    10: ('Windows 8.1', '官方已不支持 Windows 7'),
    11: ('Windows 8.1', '官方已不支持 Windows 7'),
    12: ('Windows 8.1', '官方已不支持 Windows 7'),
    13: ('Windows 8.1', '官方已不支持 Windows 7'),
    14: ('Windows 8.1', '官方已不支持 Windows 7'),
}


class _Buf(object):
    """按偏移读取字节的适配器: 内存里的包内文件。"""

    def __init__(self, data):
        self.data = data

    def read_at(self, off, n):
        return self.data[off:off + n]


class _File(object):
    """按偏移读取字节的适配器: 磁盘上的大文件(不整读进内存)。"""

    def __init__(self, fh):
        self.fh = fh

    def read_at(self, off, n):
        self.fh.seek(off)
        return self.fh.read(n)


def _cstr(src, off):
    """读 NUL 结尾的字符串(从映射后的文件偏移)。"""
    raw = src.read_at(off, 260)
    end = raw.find(b'\x00')
    if end >= 0:
        raw = raw[:end]
    try:
        return raw.decode('latin-1')
    except UnicodeDecodeError:                  # 偏移算错时不要抛, 当没读到
        return ''


def _thunk_names(src, off, ptr_size, rva2off):
    """走一遍 thunk 数组(ILT), 读出一个 DLL 被导入的每个函数名(或 #序号)。"""
    names = []
    ordinal_flag = 1 << (ptr_size * 8 - 1)
    while off:
        raw = src.read_at(off, ptr_size)
        if len(raw) < ptr_size:
            break
        val = int.from_bytes(raw, 'little')
        if not val:
            break
        if val & ordinal_flag:
            names.append('#%d' % (val & 0xFFFF))
        else:
            name_off = rva2off(val)
            if name_off is None:
                break
            raw_name = src.read_at(name_off + 2, 512)       # 跳过 2 字节 hint
            end = raw_name.find(b'\x00')
            if end >= 0:
                raw_name = raw_name[:end]
            names.append(raw_name.decode('latin-1', 'replace'))
        off += ptr_size
        if len(names) > 4096:
            break
    return names


def _imports(src, with_functions=False):
    """解析 PE 导入表(含延迟导入)。

    with_functions=False -> 返回导入的 DLL 名列表(可能重复);
    with_functions=True  -> 返回 [(DLL 名, [函数名/序号, ...]), ...]。
    非 PE 返回 None。
    """
    mz = src.read_at(0, 0x40)
    if len(mz) < 0x40 or mz[:2] != b'MZ':
        return None
    e_lfanew, = struct.unpack_from('<I', mz, 0x3C)
    coff = src.read_at(e_lfanew, 24)
    if len(coff) < 24 or coff[:4] != b'PE\x00\x00':
        return None
    nsec, = struct.unpack_from('<H', coff, 6)
    opt_size, = struct.unpack_from('<H', coff, 20)
    opt = src.read_at(e_lfanew + 24, opt_size)
    if len(opt) < opt_size or opt_size < 2:
        return None
    magic, = struct.unpack_from('<H', opt, 0)
    dd = 96 if magic == 0x10B else 112           # 数据目录起点(PE32 / PE32+)
    ptr_size = 4 if magic == 0x10B else 8
    if len(opt) < dd + 14 * 8:
        return None
    sdata = src.read_at(e_lfanew + 24 + opt_size, nsec * 40)
    sections = []
    for i in range(nsec):
        off = i * 40
        if len(sdata) < off + 24:
            break
        vsize, vaddr, rawsize, rawptr = struct.unpack_from('<IIII', sdata, off + 8)
        sections.append((vaddr, vsize or rawsize, rawptr, rawsize))

    def rva2off(rva):
        return _rva2off(sections, rva)

    # 目录 1 = 导入表(20 字节/项, Name 第 4 个 DWORD, ILT 第 1 个 / IAT 第 5 个);
    # 目录 13 = 延迟导入表(32 字节/项, rvaDLLName 第 2 个, rvaINT 第 5 个 / rvaIAT 第 4 个)
    entries = []
    for index, fields_n, name_idx, ilt_idx, iat_idx in ((1, 5, 3, 0, 4), (13, 8, 1, 4, 3)):
        rva, size = struct.unpack_from('<II', opt, dd + index * 8)
        if not rva:
            continue
        step = fields_n * 4
        off = rva2off(rva)
        limit = (size // step + 1) if size else 4096
        seen = 0
        while off and seen <= limit:
            desc = src.read_at(off, step)
            if len(desc) < step:
                break
            fields = struct.unpack_from('<%dI' % fields_n, desc, 0)
            if not any(fields):
                break
            name_off = rva2off(fields[name_idx]) if fields[name_idx] else None
            if name_off is None:
                break
            dll = _cstr(src, name_off)
            if dll:
                funcs = []
                if with_functions:
                    # 只走 ILT(OriginalFirstThunk); 被绑定过的镜像 ILT 可能为 0,
                    # 此时 IAT 里已是解析后的地址, 读不出函数名 —— 宁可空着也不瞎读
                    thunk = fields[ilt_idx]
                    funcs = _thunk_names(src, rva2off(thunk) if thunk else None,
                                         ptr_size, rva2off)
                entries.append((dll, funcs))
            off += step
            seen += 1
    if not with_functions:
        return [dll for dll, _funcs in entries]
    return entries


def _rva2off(sections, rva):
    for vaddr, vsize, rawptr, rawsize in sections:
        if vaddr <= rva < vaddr + vsize and rawptr:
            delta = rva - vaddr
            return rawptr + delta if delta < max(rawsize, vsize) else None
    return None


def pe_imports(data):
    """解析内存里的 PE 字节流, 返回导入的 DLL 名列表(非 PE 返回 None)。"""
    return _imports(_Buf(data))


def pe_imports_file(path):
    """解析磁盘上的 PE 文件, 返回导入的 DLL 名列表(非 PE 返回 None)。"""
    with open(path, 'rb') as fh:
        return _imports(_File(fh))


def pe_import_functions(path):
    """解析磁盘上的 PE 文件, 返回 [(DLL 名, [函数名/序号, ...]), ...]。

    用于「目标机上到底缺哪个导出函数」这类排障(Windows 只会说
    「找不到指定的程序」, 不告诉你是哪个 DLL 的哪个函数)。
    """
    with open(path, 'rb') as fh:
        return _imports(_File(fh), with_functions=True)


def pe_import_functions_bytes(data):
    """同 pe_import_functions, 但输入是内存里的 bytes。"""
    return _imports(_Buf(data), with_functions=True)


# ---------------------------------------------------------------- 打包辅助
# VC++ 运行库: 老系统(只装过 VC2015/2017 运行库)的 System32\MSVCP140.dll 太旧,
# 缺 onnxruntime 等新编译二进制需要的符号, 会报「找不到指定的程序」。
RUNTIME_DLLS = ('msvcp140.dll', 'vcruntime140.dll', 'vcruntime140_1.dll')


WIN7_SAFE_MAX = (14, 40)                # >= 14.40 的 VC 运行库在 Win7 上会报「找不到指定的程序」


def file_version(path):
    """用 ctypes 读 DLL/EXE 的文件版本, 返回 (major, minor, build) 或 None。"""
    import ctypes
    try:
        version = ctypes.WinDLL('version')
        size = version.GetFileVersionInfoSizeW(path, None)
        if not size:
            return None
        buf = ctypes.create_string_buffer(size)
        if not version.GetFileVersionInfoW(path, 0, size, buf):
            return None
        val = ctypes.c_void_p()
        length = ctypes.c_uint()
        if not version.VerQueryValueW(buf, '\\', ctypes.byref(val),
                                      ctypes.byref(length)):
            return None
        data = ctypes.cast(val, ctypes.POINTER(ctypes.c_uint32 * 4)).contents
        # VS_FIXEDFILEINFO: [0]=签名 [1]=结构版本 [2]=FileVersionMS [3]=FileVersionLS
        ms, ls = data[2], data[3]
        return (ms >> 16, ms & 0xFFFF, ls >> 16)
    except Exception:                       # noqa: BLE001 - 读不到版本不算错误
        return None


def _candidate_dirs(extra_dirs):
    """可能藏着 VC 运行库的目录(extra_dirs 优先, 含 wheel 自带副本如 Shapely.libs)。"""
    dirs = []
    for extra in extra_dirs:
        if extra:
            dirs.append(str(extra))
    try:
        import site
        bases = list(site.getsitepackages()) + [site.getusersitepackages()]
        for base in bases:
            if not base or not os.path.isdir(base):
                continue
            dirs.append(base)
            try:
                for name in os.listdir(base):
                    sub = os.path.join(base, name)
                    if name.endswith('.libs') and os.path.isdir(sub):
                        dirs.append(sub)
            except OSError:
                pass
    except Exception:                       # noqa: BLE001
        pass
    dirs.append(os.path.dirname(sys.executable))
    dirs.append(os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32'))
    return dirs


def _better(ver, old):
    """版本比较: 先看是不是 Win7 安全区(< 14.40), 再看谁更新。"""
    def rank(v):
        if v is None:
            return (0, (0, 0, 0))
        return (2 if v < WIN7_SAFE_MAX else 1, v)
    return rank(ver) > rank(old)


def runtime_dll_sources(verbose=True, extra_dirs=()):
    """挑一套「Win7 也能用」的 VC 运行库, 返回 {小写名: 路径}。

    优先 VS2019 时代版本(< 14.40, 官方支持 Win7 目标): 先看 extra_dirs(通常是
    项目内自带的 ``_runtime\\vc-win7-x64``), 再看 wheel 自带副本
    (``site-packages\\Shapely.libs\\msvcp140-<hash>.dll`` = 14.29), 最后才是
    Python 安装目录与 System32(14.4x/14.5x 在 Win7 上会「找不到指定的程序」)。
    """
    found = {}                              # 名字 -> (版本, 路径, 来源目录)
    for folder in _candidate_dirs(extra_dirs):
        try:
            entries = os.listdir(folder)
        except OSError:
            continue
        for fn in entries:
            low = fn.lower()
            if not low.endswith('.dll'):
                continue
            for want in RUNTIME_DLLS:
                stem = want[:-4]                # msvcp140 / vcruntime140 / vcruntime140_1
                if not low.startswith(stem):
                    continue
                tail = low[len(stem):-4]
                # 允许: 原名; 或 delvewheel 的私藏副本 msvcp140-<32位hash>.dll
                # 排除: _1/_2/_atomic_wait/_codecvt_ids/_threads/_clr0400 与 debug 版(d)
                if not (tail == '' or re.match(r'^-[0-9a-f]{16,}$', tail)):
                    continue
                ver = file_version(os.path.join(folder, fn))
                old = found.get(want)
                if old is None or _better(ver, old[0]):
                    found[want] = (ver, os.path.join(folder, fn), folder)
    if verbose:
        for want in RUNTIME_DLLS:
            if want in found:
                ver, path, _f = found[want]
                print('[*] 选用 %s %s -> %s' % (
                    want, '.'.join(str(x) for x in (ver or ('?',))), path))
            else:
                print('[!] 没找到 %s, 老系统可能报「找不到指定的程序」' % want)
    sources = {k: v[1] for k, v in found.items()}
    if any(re.search(r'-[0-9a-fA-F]{16,}\.dll$', os.path.basename(p))
           for p in sources.values()):
        # wheel 私藏副本内部 import 的是改名后的兄弟 DLL, 必须先「去改名」再分发
        tmp = tempfile.mkdtemp(prefix='paofen_rt_')
        sources = {k: demangle_runtime_copy(p, tmp, verbose)
                   for k, p in sources.items()}
    return sources


def verify_runtime_covers(toc, sources, verbose=True):
    """校验「挑出的运行库」是否覆盖包内二进制对它的全部引用。

    返回缺失清单 [(二进制, 运行库名, 函数)]; sources 传 {小写名: 路径}。
    这样打包时就能确认: 发出去的那份 MSVCP140/VCRUNTIME140 真的够用(实盘事故就是这个)。
    """
    import ctypes
    k32 = ctypes.WinDLL('kernel32')
    k32.GetProcAddress.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    k32.GetProcAddress.restype = ctypes.c_void_p
    handles = {}
    missing = []
    for dest, src, _typ in list(toc):
        if not str(dest).lower().endswith(('.pyd', '.dll')):
            continue
        try:
            entries = pe_import_functions(src) or []
        except Exception:                   # noqa: BLE001 - 打包期诊断
            continue
        for dll, funcs in entries:
            key = dll.lower()
            if key not in sources:
                continue
            if key not in handles:
                try:
                    handles[key] = ctypes.WinDLL(sources[key])
                except OSError as e:
                    handles[key] = None
                    missing.append((str(dest), dll, '(运行库自身加载失败: %s)' % e))
            handle = handles[key]
            if handle is None:
                continue
            for fn in funcs:
                if fn.startswith('#'):
                    continue
                if not k32.GetProcAddress(handle._handle,
                                          fn.encode('ascii', 'replace')):
                    missing.append((str(dest), dll, fn))
    if verbose:
        if missing:
            print('[X] 运行库覆盖不全 %d 条:' % len(missing))
            for dest, dll, fn in missing[:20]:
                print('     %s 需要 %s!%s' % (dest, dll, fn))
        else:
            print('[*] 运行库覆盖校验: 包内二进制需要的符号全部齐备')
    return missing


def add_runtime_beside_binaries(toc, verbose=True, extra_dirs=(), require_win7=False):
    """把 VC++ 运行库换成 Win7 安全版, 并补到「引用了它的」二进制所在目录。

    做四件事:
      1) 挑一套 Win7 能用的运行库(优先 extra_dirs 里自带的 ``_runtime\\vc-win7-x64``,
         其次 wheel 私藏副本 = VS2019 的 14.2x);
      2) 把 toc 里已有的 MSVCP140/VCRUNTIME140* 换成它(PyInstaller 默认从 System32 收);
      3) 子目录里的扩展(如 onnxruntime\\capi\\*.pyd)加载时 Windows 先找「它自己所在目录」
         再找 system32 —— 包根目录那份轮不到, 于是用了目标机 system32 的旧版, 报
         「DLL load failed ... 找不到指定的程序」; 所以谁引用就补到它旁边;
      4) 校验这套运行库能否覆盖包内全部引用; require_win7=True 时版本过高/覆盖不全直接
         抛错(宁可打包失败, 也别再发出对方起不来的包)。
    toc 传 PyInstaller 的 a.binaries。
    """
    sources = runtime_dll_sources(verbose=verbose, extra_dirs=extra_dirs)
    if require_win7:
        bad = []
        for name in RUNTIME_DLLS:
            ver = file_version(sources[name]) if name in sources else None
            if ver is None or ver >= WIN7_SAFE_MAX:
                bad.append('%s=%s' % (name, ver))
        if bad:
            raise RuntimeError(
                'Win7 版必须用 VS2019(<14.40)的 VC 运行库, 当前选中: %s\n'
                '请把 VS2019 的 msvcp140.dll / vcruntime140.dll / vcruntime140_1.dll '
                '放进 _runtime\\vc-win7-x64\\ 再打包(System32 里的 14.4x/14.5x 在 Win7 上'
                '会报「找不到指定的程序」)。' % ', '.join(bad))
    replaced = 0
    for i, (dest, src, typ) in enumerate(list(toc)):
        base = os.path.basename(str(dest)).lower()
        if base in sources and os.path.normcase(str(src)) != os.path.normcase(sources[base]):
            toc[i] = (dest, sources[base], typ)
            replaced += 1
    needed = {}
    for dest, src, _typ in list(toc):
        if not str(dest).lower().endswith(('.pyd', '.dll')):
            continue
        try:
            names = {str(n).lower() for n in (pe_imports_file(src) or [])}
        except Exception:                   # noqa: BLE001 - 打包期诊断, 别中断构建
            continue
        hit = names & set(RUNTIME_DLLS)
        if hit:
            needed.setdefault(os.path.dirname(str(dest)), set()).update(hit)
    have = {str(d).lower() for d, _s, _t in toc}
    added = []
    for folder, names in sorted(needed.items()):
        for name in sorted(names):
            if name not in sources:
                continue
            dest = os.path.join(folder, name) if folder else name
            if dest.lower() in have:
                continue
            toc.append((dest, sources[name], 'BINARY'))
            have.add(dest.lower())
            added.append(dest)
    if verbose:
        print('[*] 替换已有运行库 %d 处, 补齐到子目录 %d 处' % (replaced, len(added)))
    missing = verify_runtime_covers(toc, sources, verbose=verbose)
    if require_win7 and missing:
        raise RuntimeError('挑出的 VC 运行库覆盖不全(见上), 这个包在目标机上会「找不到指定的程序」')
    return added


_MANGLED_RE = re.compile(rb'([A-Za-z0-9_]+?)-[0-9a-fA-F]{16,}\.dll')


def demangle_runtime_copy(src, dest_dir, verbose=True):
    """把 wheel 私藏副本(msvcp140-<hash>.dll) 变成可直接分发、可加载的 msvcp140.dll。

    这类副本内部还会 import 兄弟文件的改名版(如 ``vcruntime140_1-<hash>.dll``), 直接
    改名分发就会「找不到指定的模块」。这里把导入表里的改名依赖就地改回标准名(用 NUL
    补齐, 不改变文件长度), 另存到 dest_dir, 返回新路径。
    """
    with open(src, 'rb') as f:
        data = bytearray(f.read())
    hits = list(_MANGLED_RE.finditer(bytes(data)))
    out_name = re.sub(r'-[0-9a-fA-F]{16,}(?=\.dll$)', '', os.path.basename(src))
    if not hits:
        return src
    for match in hits:
        base = match.group(1) + b'.dll'
        start, end = match.start(), match.end()
        if end - start < len(base):
            continue
        data[start:start + len(base)] = base
        for i in range(start + len(base), end):
            data[i] = 0
    if not os.path.isdir(dest_dir):
        os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, out_name)
    with open(dest, 'wb') as f:
        f.write(bytes(data))
    if verbose:
        print('[*] 去改名副本: %s (%d 处内部依赖改回标准名) -> %s'
              % (os.path.basename(src), len(hits), dest))
    return dest


def probe_missing_exports(files, limit=80, extra_dirs=()):
    """按加载器搜索顺序, 逐个查每个导入函数在「当前机器」上是否存在。

    对文件 F 的每个导入 (DLL, 函数): 先看 F 同目录有没有该 DLL(加载器就是这么找的),
    有就用它, 否则按名字加载(系统路径); 再 GetProcAddress 问函数在不在。
    返回 [(文件, DLL, 函数, 实际加载的 DLL 路径), ...] —— Windows 只会含糊地报
    「找不到指定的程序」, 这样能直接指出是哪个 DLL 的哪个函数。
    """
    import ctypes
    k32 = ctypes.WinDLL('kernel32')
    k32.GetProcAddress.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    k32.GetProcAddress.restype = ctypes.c_void_p
    k32.GetModuleFileNameW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint]
    k32.GetModuleFileNameW.restype = ctypes.c_uint
    handles = {}
    missing = []
    for path in files:
        for dll, funcs in (pe_import_functions(path) or []):
            near = os.path.join(os.path.dirname(path), dll)
            use = near
            if not os.path.isfile(use):             # 加载器还会看 .libs 目录(如 delvewheel)
                for d in extra_dirs:
                    cand = os.path.join(d, dll)
                    if os.path.isfile(cand):
                        use = cand
                        break
            if not os.path.isfile(use):
                use = dll                           # 交给系统搜索路径
            key = os.path.normcase(use)
            if key not in handles:
                try:
                    handles[key] = ctypes.WinDLL(use)
                except OSError as e:
                    handles[key] = None
                    missing.append((path, dll, '(整个 DLL 加载失败)', str(e)))
            handle = handles[key]
            if handle is None:
                continue
            buf = ctypes.create_unicode_buffer(512)
            k32.GetModuleFileNameW(handle._handle, buf, 512)
            for fn in funcs:
                if fn.startswith('#'):          # 序号导入不查
                    continue
                if not k32.GetProcAddress(handle._handle,
                                          fn.encode('ascii', 'replace')):
                    missing.append((path, dll, fn, buf.value))
                    if len(missing) > limit:
                        return missing
    return missing


def apiset_hits(names):
    """挑出 Windows 7 缺的 API 集(返回列表, 不含 UCRT)。"""
    return sorted({n for n in names
                   if APISET_RE.match(n) and not n.lower().startswith(UCRT_PREFIX)})


def ucrt_hits(names):
    return sorted({n for n in names if n.lower().startswith(UCRT_PREFIX)})


def _iter_dir(root):
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if fn.lower().endswith(('.dll', '.pyd', '.exe')):
                yield fn, os.path.join(dirpath, fn)


def inspect(path):
    """分析一个 exe(onefile) 或 dist 目录(onedir), 返回统计结果。"""
    info = {'path': path, 'mode': '', 'count': 0, 'py_dll': None,
            'bundled_ucrt': [], 'blockers': {}, 'ucrt_users': {}, 'errors': []}

    def feed(name, names):
        if names is None:                       # 不是 PE(比如纯资源 dll), 跳过
            return
        info['count'] += 1
        if re.match(r'^python3\d+\.dll$', name, re.I):
            info['py_dll'] = name.lower()
        if name.lower().startswith(('api-ms-win', 'ext-ms-win')) or \
                name.lower() == 'ucrtbase.dll':
            info['bundled_ucrt'].append(name)
        bad = apiset_hits(names)
        if bad:
            info['blockers'][name] = bad
        if ucrt_hits(names):
            info['ucrt_users'][name] = ucrt_hits(names)

    if os.path.isdir(path):
        info['mode'] = 'onedir 文件夹'
        for name, full in _iter_dir(path):
            try:
                feed(name, pe_imports_file(full))
            except OSError as e:
                info['errors'].append('%s: %s' % (name, e))
    elif os.path.isfile(path):
        info['mode'] = 'onefile 单文件'
        feed(os.path.basename(path), pe_imports_file(path))
        try:                                    # 逐个取出包内 dll/pyd 看导入表
            from PyInstaller.archive.readers import CArchiveReader
            reader = CArchiveReader(path)
            for name in sorted(reader.toc):
                if name.lower().endswith(('.dll', '.pyd', '.exe')):
                    try:
                        feed(name, pe_imports(reader.extract(name)))
                    except Exception as e:      # noqa: BLE001 - 单个条目失败不影响整体
                        info['errors'].append('%s: %s' % (name, e))
        except ImportError:
            info['errors'].append('未装 PyInstaller, 无法展开单文件包内容(只查了 exe 自身)')
        except Exception as e:                  # noqa: BLE001 - 不是 PyInstaller 包
            info['errors'].append('不是 PyInstaller 单文件包? %s' % e)
    else:
        info['errors'].append('文件不存在')
    return info


def py_min_os(py_dll):
    """从包内 python3.x dll 名字推官方最低系统要求。"""
    if not py_dll:
        return None, '未找到包内 python3x.dll(可能不是 PyInstaller 包)'
    m = re.match(r'^python3(\d+)\.dll$', py_dll)
    minor = int(m.group(1)) if m else 0
    os_name, note = PY_MIN_OS.get(minor, ('未知', ''))
    return os_name, 'Python 3.%d' % minor + (' —— ' + note if note else '')


def report(info, show_all=False):
    """生成人类可读的检查报告。"""
    lines = []
    add = lines.append
    path = info['path']
    size = os.path.getsize(path) / 1024.0 / 1024.0 if os.path.isfile(path) else 0
    add('== %s  (%.1f MB, %s)' % (path, size, info['mode']))
    add('   包内可执行文件: %d 个' % info['count'])

    min_os, py_desc = py_min_os(info['py_dll'])
    add('   打包运行时: %s -> %s' % (info['py_dll'] or '(未知)', py_desc))

    bundled_apis = sorted(set(info['bundled_ucrt']))
    if bundled_apis:
        add('   包内自带 API 集/UCRT 文件: %s' % ', '.join(bundled_apis))
    else:
        add('   包内未自带 UCRT(api-ms-win-crt-*.dll / ucrtbase.dll)'
            ' —— Windows 7 需先装 VC++ 2015-2022 运行库(KB2999226)')

    for e in info['errors']:
        add('   [!] %s' % e)

    if show_all and info['ucrt_users']:
        add('   -- 引用 UCRT 的文件(%d): %s' % (
            len(info['ucrt_users']), ', '.join(sorted(info['ucrt_users']))[:400]))

    add('')
    if info['blockers']:
        add('   [X] 无法在 Windows 7 启动: 下列二进制引用了 Win8+ 才有的 API 集,')
        add('       Windows 7 上会直接弹「无法启动此程序, 因为计算机中丢失 xxx.dll」:')
        for name in sorted(info['blockers']):
            add('         %-34s -> %s' % (name, ', '.join(info['blockers'][name])))
        add('')
        add('   处理办法(任选其一):')
        add('     1) 老系统/不确定对方系统时, 用 Windows 7 兼容版打包:')
        add('        .\\build_exe_win7.ps1        (Python 3.8 + 老依赖, 不引用这些 API 集)')
        add('     2) 目标机是 Windows 7 且必须用本包: 让对方装 KB2533623(补 API 集),')
        add('        再装 VC++ 2015-2022 运行库(UCRT); 装完用 --selftest 复验。')
        add('     3) 目标机是 Windows 8.1/10/11: 本包可用, 无需处理。')
    else:
        add('   [OK] 没有引用 Windows 8+ 专有 API 集, Windows 7 可以启动。')
        if info['ucrt_users'] and not bundled_apis:
            add('        (Windows 7 上仍需 UCRT: 装 VC++ 2015-2022 运行库/KB2999226,')
            add('         否则会报「丢失 api-ms-win-crt-runtime-l1-1-0.dll」)')
    add('   官方系统要求(Python 运行时): %s' % min_os)
    add('')
    return '\n'.join(lines)


def main(argv):
    args = [a for a in argv[1:] if not a.startswith('-')]
    show_all = any(a == '--all' for a in argv[1:])
    if not args:
        here = os.path.dirname(os.path.abspath(__file__))
        args = [os.path.join(here, 'dist', '跑分绩效汇总自动生成工具.exe')]
    blocked = False
    for path in args:
        info = inspect(path)
        text = report(info, show_all)
        print(text)
        if not info['count'] and info['errors']:
            return 2
        blocked = blocked or bool(info['blockers'])
    print('结论: %s' % ('发现 Windows 7 起不来的依赖, 建议按上面的办法处理'
                       if blocked else 'Windows 7 可启动(注意 UCRT 前提)'))
    return 1 if blocked else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
