#!/usr/bin/env python3
"""report.py — Generador de reportes de sesión de docflow.

Lee el results.tsv de una sesión y produce reportes HTML, CSV, JSON o
Markdown con estadísticas: totales por estado, porcentaje de éxito, tiempos
(total/promedio), tamaños de entrada/salida y desglose por formato.

Uso:
    report.py --session DIR --format html [--out-dir DIR] [--stdout]
"""

import argparse
import csv
import html
import io
import json
import sys
from datetime import datetime
from pathlib import Path

COLUMNS = ["status", "task", "input", "output", "ms", "bytes_in", "bytes_out", "message"]


def load_session(session_dir: Path) -> dict:
    results = []
    tsv = session_dir / "results.tsv"
    if tsv.is_file():
        for line in tsv.read_text(encoding="utf-8", errors="replace").splitlines():
            parts = line.split("\t")
            parts += [""] * (len(COLUMNS) - len(parts))
            row = dict(zip(COLUMNS, parts))
            for key in ("ms", "bytes_in", "bytes_out"):
                try:
                    row[key] = int(row[key])
                except ValueError:
                    row[key] = 0
            results.append(row)
    task_file = session_dir / "task"
    task = task_file.read_text().strip() if task_file.is_file() else "?"
    return {"session": session_dir.name, "task": task, "results": results}


def compute_stats(data: dict) -> dict:
    results = data["results"]
    by_status: dict = {}
    by_format: dict = {}
    for row in results:
        by_status[row["status"]] = by_status.get(row["status"], 0) + 1
        ext = Path(row["input"]).suffix.lstrip(".").lower() or "?"
        fmt = by_format.setdefault(ext, {"total": 0, "ok": 0, "ms": 0})
        fmt["total"] += 1
        if row["status"] == "ok":
            fmt["ok"] += 1
            fmt["ms"] += row["ms"]

    converted = [r for r in results if r["status"] == "ok"]
    attempted = [r for r in results if r["status"] in ("ok", "fail", "protected")]
    total_ms = sum(r["ms"] for r in converted)
    stats = {
        "total": len(results),
        "by_status": by_status,
        "success_pct": round(100 * len(converted) / len(attempted), 1) if attempted else 100.0,
        "total_ms": total_ms,
        "avg_ms": round(total_ms / len(converted)) if converted else 0,
        "bytes_in": sum(r["bytes_in"] for r in converted),
        "bytes_out": sum(r["bytes_out"] for r in converted),
        "by_format": by_format,
    }
    return stats


def human(n: int) -> str:
    for unit in ("B", "KiB", "MiB", "GiB"):
        if abs(n) < 1024 or unit == "GiB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024
    return f"{n}"


def render_json(data: dict, stats: dict) -> str:
    return json.dumps({"session": data["session"], "task": data["task"],
                       "stats": stats, "results": data["results"]},
                      ensure_ascii=False, indent=2)


def render_csv(data: dict, stats: dict) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(COLUMNS)
    for row in data["results"]:
        writer.writerow([row[c] for c in COLUMNS])
    return buf.getvalue()


def render_md(data: dict, stats: dict) -> str:
    lines = [
        f"# Reporte docflow — sesión {data['session']}",
        "",
        f"- **Tarea:** {data['task']}",
        f"- **Archivos:** {stats['total']}",
        f"- **Éxito:** {stats['success_pct']}%",
        f"- **Tiempo total:** {stats['total_ms'] / 1000:.1f} s (promedio {stats['avg_ms']} ms/archivo)",
        f"- **Tamaño:** {human(stats['bytes_in'])} → {human(stats['bytes_out'])}",
        "",
        "## Por estado",
        "",
        "| Estado | Archivos |",
        "|---|---|",
    ]
    lines += [f"| {s} | {n} |" for s, n in sorted(stats["by_status"].items())]
    lines += ["", "## Por formato", "", "| Formato | Total | OK | ms promedio |", "|---|---|---|---|"]
    for ext, f in sorted(stats["by_format"].items()):
        avg = round(f["ms"] / f["ok"]) if f["ok"] else "-"
        lines.append(f"| {ext} | {f['total']} | {f['ok']} | {avg} |")
    failures = [r for r in data["results"] if r["status"] in ("fail", "protected")]
    if failures:
        lines += ["", "## Errores", "", "| Archivo | Estado | Detalle |", "|---|---|---|"]
        lines += [f"| {r['input']} | {r['status']} | {r['message']} |" for r in failures]
    return "\n".join(lines) + "\n"


def render_html(data: dict, stats: dict) -> str:
    def row_class(status):
        return {"ok": "ok", "fail": "fail", "protected": "fail"}.get(status, "muted")

    rows = "\n".join(
        f"<tr class='{row_class(r['status'])}'><td>{html.escape(r['status'])}</td>"
        f"<td>{html.escape(r['input'])}</td><td>{html.escape(r['output'])}</td>"
        f"<td>{r['ms']}</td><td>{human(r['bytes_in'])}</td>"
        f"<td>{human(r['bytes_out'])}</td><td>{html.escape(r['message'])}</td></tr>"
        for r in data["results"])
    fmt_rows = "\n".join(
        f"<tr><td>{html.escape(ext)}</td><td>{f['total']}</td><td>{f['ok']}</td>"
        f"<td>{round(f['ms'] / f['ok']) if f['ok'] else '-'}</td></tr>"
        for ext, f in sorted(stats["by_format"].items()))
    status_rows = "\n".join(
        f"<tr><td>{html.escape(s)}</td><td>{n}</td></tr>"
        for s, n in sorted(stats["by_status"].items()))

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<title>docflow — reporte {html.escape(data['session'])}</title>
<style>
 body {{ font-family: system-ui, sans-serif; margin: 2rem auto; max-width: 72rem; color: #222; }}
 h1 {{ border-bottom: 2px solid #4a7; padding-bottom: .3rem; }}
 table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; font-size: .9rem; }}
 th, td {{ border: 1px solid #ccc; padding: .35rem .6rem; text-align: left; }}
 th {{ background: #f0f4f2; }}
 tr.ok td:first-child {{ color: #187a3c; font-weight: 600; }}
 tr.fail td:first-child {{ color: #b3261e; font-weight: 600; }}
 tr.muted td:first-child {{ color: #777; }}
 .cards {{ display: flex; gap: 1rem; flex-wrap: wrap; }}
 .card {{ background: #f6f8f7; border: 1px solid #dde; border-radius: .5rem; padding: .8rem 1.2rem; }}
 .card b {{ display: block; font-size: 1.4rem; }}
</style></head><body>
<h1>docflow — reporte de conversión</h1>
<p>Sesión <code>{html.escape(data['session'])}</code> · tarea <code>{html.escape(data['task'])}</code>
 · generado {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
<div class="cards">
 <div class="card"><b>{stats['total']}</b> archivos</div>
 <div class="card"><b>{stats['success_pct']}%</b> éxito</div>
 <div class="card"><b>{stats['total_ms'] / 1000:.1f} s</b> tiempo total</div>
 <div class="card"><b>{stats['avg_ms']} ms</b> promedio/archivo</div>
 <div class="card"><b>{human(stats['bytes_in'])} → {human(stats['bytes_out'])}</b> tamaño</div>
</div>
<h2>Por estado</h2>
<table><tr><th>Estado</th><th>Archivos</th></tr>{status_rows}</table>
<h2>Por formato</h2>
<table><tr><th>Formato</th><th>Total</th><th>OK</th><th>ms promedio</th></tr>{fmt_rows}</table>
<h2>Detalle</h2>
<table><tr><th>Estado</th><th>Entrada</th><th>Salida</th><th>ms</th><th>Entrada</th><th>Salida</th><th>Mensaje</th></tr>
{rows}</table>
</body></html>
"""


RENDERERS = {"json": render_json, "csv": render_csv, "md": render_md, "html": render_html}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--session", required=True, type=Path)
    ap.add_argument("--format", required=True, choices=sorted(RENDERERS))
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--stdout", action="store_true")
    args = ap.parse_args()

    data = load_session(args.session)
    stats = compute_stats(data)
    output = RENDERERS[args.format](data, stats)

    if args.stdout or args.out_dir is None:
        sys.stdout.write(output)
        return 0

    args.out_dir.mkdir(parents=True, exist_ok=True)
    dest = args.out_dir / f"reporte_{data['session']}.{args.format}"
    dest.write_text(output, encoding="utf-8")
    print(dest)
    return 0


if __name__ == "__main__":
    sys.exit(main())
