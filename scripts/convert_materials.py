"""批量转换测试素材：tests/materials/*.pdf -> outputs/<素材名>/。

用法::

    python scripts/convert_materials.py                 # 转换全部素材
    python scripts/convert_materials.py 测试1.pdf       # 只转换指定文件
    python scripts/convert_materials.py --no-ocr       # 关闭 OCR（仅结构化路径）

每个素材输出：
    outputs/<素材名>/<素材名>.pptx   转换结果
    outputs/<素材名>/report.json     可编辑性验证报告
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.services.converter import ConvertOptions, convert_pdf  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
MATERIALS_DIR = ROOT / "tests" / "materials"
OUTPUT_ROOT = ROOT / "outputs"


def main() -> int:
    parser = argparse.ArgumentParser(description="批量转换 PDF 测试素材")
    parser.add_argument("names", nargs="*", help="要转换的文件名（可省略，默认全部）")
    parser.add_argument("--no-ocr", action="store_true", help="关闭 OCR 文字还原")
    parser.add_argument("--min-confidence", type=float, default=0.5, help="OCR 置信度阈值")
    parser.add_argument("--output", default=str(OUTPUT_ROOT), help="输出根目录")
    args = parser.parse_args()

    if not MATERIALS_DIR.exists():
        print(f"素材目录不存在：{MATERIALS_DIR}")
        return 1

    pdfs = sorted(MATERIALS_DIR.glob("*.pdf"))
    if args.names:
        wanted = set(args.names)
        pdfs = [item for item in pdfs if item.name in wanted]
    if not pdfs:
        print("没有找到待转换的 PDF")
        return 1

    options = ConvertOptions(
        enable_ocr=not args.no_ocr,
        ocr_min_confidence=args.min_confidence,
    )
    output_root = Path(args.output)
    summary: list[dict] = []

    for pdf in pdfs:
        target = output_root / pdf.stem
        print(f"\n=== 转换 {pdf.name} ===")
        started = time.perf_counter()
        try:
            result = convert_pdf(pdf, target, pdf.name, options)
        except Exception as exc:
            print(f"  失败：{exc}")
            summary.append({"name": pdf.name, "status": "failed", "error": str(exc)})
            continue
        elapsed = time.perf_counter() - started
        info = result.report["summary"]
        print(
            f"  完成：{result.output_path.name}  用时 {elapsed:.1f}s  "
            f"页数 {result.report['page_count']}  "
            f"文本行 {info['text_lines']}  图片 {info['images']}  "
            f"可编辑对象 {info['editable_objects']}  "
            f"可编辑占比 {info['editable_ratio']:.2%}"
        )
        summary.append(
            {
                "name": pdf.name,
                "status": "ok",
                "output": str(result.output_path.relative_to(ROOT)),
                "report": str(result.report_path.relative_to(ROOT)),
                "elapsed_seconds": round(elapsed, 2),
                "summary": info,
            }
        )

    index_path = output_root / "conversion_summary.json"
    output_root.mkdir(parents=True, exist_ok=True)
    index_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n汇总已写入 {index_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
