#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
_paths.py —— 脚本层 realpath 路径校验（纯标准库，供 scripts/*.py 共用，D5.4）

策略：
  - 输入/输出路径先 expanduser → realpath，消除相对路径、``..``、符号链接带来的歧义；
  - 输入：必须存在、必须是常规文件；拒绝符号链接与 Windows 重解析点（防止借链接
    逃逸到预期目录之外）；
  - 输出：父目录必须存在；目标不得是目录，也不得是符号链接；
  - 校验失败抛 PathError，调用方应以退出码 2 结束并打印可读错误。

用法（在 scripts/ 下的脚本中）：
    from _paths import PathError, resolve_input, resolve_output
    src = resolve_input(args.input)
    out = resolve_output(args.out)
"""

import os


class PathError(ValueError):
    """路径校验失败。"""


def _is_link(path):
    """符号链接 / Windows 重解析点（junction、mount point 等）。"""
    if os.path.islink(path):
        return True
    try:
        st = os.lstat(path)
        return bool(getattr(st, "st_reparse_tag", 0))
    except OSError:
        return False


def _canonical(raw):
    try:
        return os.path.realpath(os.path.expanduser(raw))
    except (ValueError, OSError) as e:
        raise PathError("路径非法: %s" % e)


def resolve_input(path):
    """校验输入文件并返回其 canonical 真实路径。"""
    raw = os.path.expanduser(str(path)) if path else ""
    if not raw:
        raise PathError("输入路径为空")
    if "\x00" in raw:
        raise PathError("输入路径含非法字符")
    if _is_link(raw):
        raise PathError("拒绝符号链接输入（真实路径 %s）: %s" % (_canonical(raw), path))
    real = _canonical(raw)
    if not os.path.exists(real):
        raise PathError("文件不存在: %s" % path)
    if not os.path.isfile(real):
        raise PathError("不是常规文件: %s（真实路径 %s）" % (path, real))
    return real


def resolve_output(path):
    """校验输出文件并返回其 canonical 真实路径（父目录须已存在）。"""
    raw = os.path.expanduser(str(path)) if path else ""
    if not raw:
        raise PathError("输出路径为空")
    if "\x00" in raw:
        raise PathError("输出路径含非法字符")
    if _is_link(raw):
        raise PathError("拒绝写入符号链接: %s" % path)
    real = _canonical(raw)
    parent = os.path.dirname(real)
    if not os.path.isdir(parent):
        raise PathError("输出目录不存在: %s" % parent)
    if os.path.isdir(real):
        raise PathError("输出路径是目录: %s" % path)
    return real
