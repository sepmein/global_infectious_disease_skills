"""真实并发 HTTP 探测候选官方机构入口（仅标准库）。

输出 JSON：每个 code 的探测结果，含真实 HTTP 状态码/最终 URL/错误类别/页面标题片段。
不写入任何用户资料，只读取候选与名单。
"""
from __future__ import annotations

import json
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
import skill_config as sc  # noqa: E402

REFS = Path(__file__).resolve().parents[2] / "references"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
CONNECT_TIMEOUT = 8.0
READ_TIMEOUT = 15.0


def load_candidates() -> dict[str, dict]:
    sheets = sc.read_workbook_sheets(REFS / "各国权威机构参考.xlsx")
    rows = sheets["国家监测数据源"]
    hdr = rows[0]
    idx = {h: i for i, h in enumerate(hdr)}
    out: dict[str, dict] = {}
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        code = r[idx["地区代码"]].strip()
        out[code] = {
            "name_zh": r[idx["中文标准名称"]].strip(),
            "org_priority": r[idx["国家优先监测机构"]].strip(),
            "candidate_url": r[idx["国家监测数据源 URL"]].strip(),
            "region_cross_org": r[idx["区域交叉来源"]].strip() if "区域交叉来源" in idx else "",
            "region_cross_url": r[idx["区域交叉来源 URL"]].strip() if "区域交叉来源 URL" in idx else "",
        }
    return out


def classify_error(exc: Exception) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        return f"HTTP {exc.code}"
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", exc)
        msg = str(reason).lower()
        if "timed out" in msg or "timeout" in msg:
            if "read" in msg:
                return "读取超时"
            return "连接超时"
        if "name or service not known" in msg or "getaddrinfo" in msg or "nodename" in msg:
            return "DNS解析失败"
        if "certificate" in msg or "ssl" in msg or "tls" in msg:
            return "证书错误"
        if "tls" in msg or "handshake" in msg:
            return "TLS错误"
        if "connection refused" in msg:
            return "连接被拒绝"
        if "connection reset" in msg:
            return "连接重置"
        if "network is unreachable" in msg or "no route" in msg:
            return "网络不可达"
        return f"连接失败:{msg[:40]}"
    if isinstance(exc, ConnectionError):
        return "连接被拒绝"
    return f"未知错误:{type(exc).__name__}"


def probe(url: str) -> dict:
    result = {
        "url": url,
        "status": None,
        "final_url": url,
        "error": None,
        "title": None,
        "redirected": False,
    }
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})
    try:
        with urllib.request.urlopen(req, timeout=CONNECT_TIMEOUT + READ_TIMEOUT) as resp:
            status = resp.getcode()
            final_url = resp.geturl()
            data = resp.read(200000)
            result["status"] = status
            result["final_url"] = final_url
            result["redirected"] = (final_url.rstrip("/") != url.rstrip("/"))
            text = data.decode("utf-8", "ignore")
            t0 = text.find("<title")
            if t0 != -1:
                t1 = text.find("</title>", t0)
                if t1 != -1:
                    title = text[t0 + 6:t1].strip()
                    title = title.replace("\n", " ").replace("\r", " ")
                    result["title"] = title[:120]
    except urllib.error.HTTPError as e:
        result["status"] = e.code
        try:
            result["final_url"] = e.geturl()
        except Exception:
            pass
        result["redirected"] = (result["final_url"].rstrip("/") != url.rstrip("/"))
    except Exception as e:  # noqa: BLE001
        result["error"] = classify_error(e)
    return result


def main() -> None:
    cands = load_candidates()
    pl = sc.load_participants()
    by_code = pl.by_code
    jobs = []
    for code, c in cands.items():
        jobs.append((code, c))
    results: dict[str, dict] = {}
    lock = threading.Lock()

    def work(code, c):
        res = probe(c["candidate_url"])
        with lock:
            results[code] = {**c, "probe": res}
        return code

    with ThreadPoolExecutor(max_workers=24) as ex:
        futs = [ex.submit(work, code, c) for code, c in jobs]
        for f in as_completed(futs):
            f.result()

    out = Path(__file__).with_name("probe_results.json")
    payload = {
        "probe_date": "2026-09-09",
        "total_candidates": len(results),
        "results": results,
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    # summary
    ok = sum(1 for v in results.values() if v["probe"]["status"] and str(v["probe"]["status"]).startswith("2"))
    blocked = sum(1 for v in results.values() if v["probe"]["status"] in (401, 403, 429, 503) or (v["probe"]["error"] and ("机器人" in v["probe"]["error"] or "bot" in v["probe"]["error"].lower())))
    err = sum(1 for v in results.values() if v["probe"]["error"])
    print(f"candidates={len(results)} http2xx={ok} blocked-ish={blocked} errors={err}")
    print("written:", out)


if __name__ == "__main__":
    main()
