"""HTTP 层端到端验证：启动服务 → 上传 → 转换 → 下载 → 查看报告。

用法::

    python scripts/e2e_http_check.py                       # 自动启动服务再测试
    python scripts/e2e_http_check.py --base http://...     # 测试已在运行的服务

脚本会在临时目录中启动一个 uvicorn 进程，使用真实素材走一遍完整链路，
最后打印每一步的结果。
"""

from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MATERIALS_DIR = ROOT / "tests" / "materials"


def _free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _request(url: str, method: str = "GET", data: bytes | None = None, headers: dict | None = None):
    request = urllib.request.Request(url, data=data, method=method, headers=headers or {})
    with urllib.request.urlopen(request, timeout=600) as response:
        return response.status, response.read(), dict(response.headers)


def _multipart(
    files: list[tuple[str, bytes]], options: dict[str, str] | None = None
) -> tuple[bytes, str]:
    boundary = uuid.uuid4().hex
    parts: list[bytes] = []
    for name, payload in files:
        parts.append(
            f"--{boundary}\r\n".encode()
            + f'Content-Disposition: form-data; name="files"; filename="{name}"\r\n'.encode()
            + b"Content-Type: application/pdf\r\n\r\n"
            + payload
            + b"\r\n"
        )
    for key, value in (options or {}).items():
        parts.append(
            f"--{boundary}\r\n".encode()
            + f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode()
            + value.encode()
            + b"\r\n"
        )
    parts.append(f"--{boundary}--\r\n".encode())
    return b"".join(parts), f"multipart/form-data; boundary={boundary}"


def run_checks(
    base: str, materials: list[Path], options: dict[str, str] | None = None
) -> int:
    failures = 0

    print("1) 健康检查")
    status, body, _ = _request(f"{base}/api/health")
    print(f"   HTTP {status}  {body.decode('utf-8')}")
    if status != 200:
        failures += 1

    print("2) 前端首页")
    status, body, headers = _request(f"{base}/")
    print(f"   HTTP {status}  Content-Type={headers.get('Content-Type')}  {len(body)} 字节")
    if status != 200 or b"<html" not in body.lower():
        failures += 1

    print("3) 批量上传并转换")
    payload, content_type = _multipart(
        [(item.name, item.read_bytes()) for item in materials], options
    )
    started = time.perf_counter()
    status, body, _ = _request(f"{base}/api/convert", "POST", payload, {"Content-Type": content_type})
    elapsed = time.perf_counter() - started
    print(f"   HTTP {status}  用时 {elapsed:.1f}s")
    result = json.loads(body)
    print(f"   成功 {len(result['files'])} 个，失败 {len(result['errors'])} 个")
    for item in result["files"]:
        stats = item["stats"]
        print(
            f"     - {item['source_name']}: {stats['page_count']} 页 / "
            f"文本框 {stats['text_lines']} / 可编辑占比 {stats['editable_ratio']:.0%}"
        )
    if not result["files"]:
        print(f"   错误：{result.get('errors')}")
        failures += 1
        return failures

    print("4) 下载转换结果")
    for item in result["files"]:
        status, body, headers = _request(f"{base}{item['download_url']}")
        print(
            f"   {item['source_name']}: HTTP {status}  {len(body)} 字节  "
            f"{headers.get('Content-Type')}"
        )
        if status != 200 or not body.startswith(b"PK"):
            failures += 1

    print("5) 查看可编辑性报告")
    for item in result["files"]:
        status, body, _ = _request(f"{base}{item['inspect_url']}")
        report = json.loads(body)
        print(
            f"   {item['source_name']}: HTTP {status}  {report['page_count']} 页  "
            f"可编辑对象 {report['summary']['editable_objects']}"
        )
        if status != 200:
            failures += 1

    print("6) 非法输入校验")
    payload, content_type = _multipart([("bad.txt", b"not a pdf")])
    try:
        _request(f"{base}/api/convert", "POST", payload, {"Content-Type": content_type})
        print("   非 PDF 文件未被拒绝")
        failures += 1
    except urllib.error.HTTPError as exc:
        print(f"   非 PDF 文件被正确拒绝：HTTP {exc.code}")

    try:
        _request(f"{base}/api/download/does-not-exist")
        print("   不存在的文件未被拒绝")
        failures += 1
    except urllib.error.HTTPError as exc:
        print(f"   不存在的文件被正确拒绝：HTTP {exc.code}")

    print("7) 请求级开关是否生效（关闭 OCR 后不应识别出文字）")
    if materials:
        payload, content_type = _multipart(
            [(materials[0].name, materials[0].read_bytes())], {"enable_ocr": "0"}
        )
        status, body, _ = _request(
            f"{base}/api/convert", "POST", payload, {"Content-Type": content_type}
        )
        payload_json = json.loads(body)
        lines = payload_json["files"][0]["stats"]["text_lines"] if payload_json["files"] else -1
        print(f"   enable_ocr=0 时文本行数：{lines}")
        if lines != 0:
            print("   OCR 开关未生效")
            failures += 1
        else:
            print("   OCR 开关已生效")

    print("8) 请求级开关是否生效（关闭图层分解后不应出现独立图片对象）")
    if materials:
        payload, content_type = _multipart(
            [(materials[0].name, materials[0].read_bytes())], {"enable_layering": "0"}
        )
        status, body, _ = _request(
            f"{base}/api/convert", "POST", payload, {"Content-Type": content_type}
        )
        payload_json = json.loads(body)
        stats = payload_json["files"][0]["stats"] if payload_json["files"] else {}
        print(
            f"   enable_layering=0：已分解 {stats.get('decomposed_pages')} 页 / "
            f"整页图 {stats.get('full_page_image_pages')} 页 / 图片 {stats.get('images')} 个"
        )
        if stats.get("decomposed_pages") != 0:
            print("   图层分解开关未生效")
            failures += 1
        else:
            print("   图层分解开关已生效")

    print("9) 请求级开关是否生效（开启图层分解后应出现独立图片对象）")
    if materials:
        payload, content_type = _multipart(
            [(materials[0].name, materials[0].read_bytes())], {"enable_layering": "1"}
        )
        status, body, _ = _request(
            f"{base}/api/convert", "POST", payload, {"Content-Type": content_type}
        )
        payload_json = json.loads(body)
        stats = payload_json["files"][0]["stats"] if payload_json["files"] else {}
        print(
            f"   enable_layering=1：已分解 {stats.get('decomposed_pages')} 页 / "
            f"整页图 {stats.get('full_page_image_pages')} 页 / 图片 {stats.get('images')} 个"
        )
        if not stats.get("decomposed_pages"):
            print("   图层分解未生效")
            failures += 1
        else:
            print("   图层分解已生效")

    print()
    print("端到端验证：" + ("全部通过" if failures == 0 else f"存在 {failures} 处失败"))
    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description="HTTP 端到端验证")
    parser.add_argument("--base", help="已运行服务的地址；不传则自动启动")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument(
        "--limit", type=int, default=0, help="只使用前 N 个素材（0 表示全部）"
    )
    args = parser.parse_args()

    materials = sorted(MATERIALS_DIR.glob("*.pdf"))
    if args.limit > 0:
        materials = materials[: args.limit]
    if not materials:
        print(f"素材目录为空：{MATERIALS_DIR}")
        return 1

    if args.base:
        return run_checks(args.base.rstrip("/"), materials)

    port = args.port or _free_port()
    base = f"http://127.0.0.1:{port}"
    env = dict(os.environ)
    env["PDF2PPTX_UPLOAD_DIR"] = str(ROOT / "uploads")
    env["PDF2PPTX_OUTPUT_DIR"] = str(ROOT / "outputs" / "_e2e")
    process = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "uvicorn",
            "app.main:app",
            "--host",
            "127.0.0.1",
            "--port",
            str(port),
        ],
        cwd=str(ROOT),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    try:
        for _ in range(120):
            try:
                _request(f"{base}/api/health")
                break
            except Exception:
                time.sleep(0.5)
        else:
            print("服务启动超时")
            return 1
        return run_checks(base, materials)
    finally:
        process.terminate()
        try:
            process.wait(timeout=10)
        except Exception:
            process.kill()


if __name__ == "__main__":
    raise SystemExit(main())
