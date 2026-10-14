"""Social Scraper — runner: port dari fungsi on_click() di notebook.

Nilai form (dulu ipywidgets) sekarang datang dari dict `P` yang dikirim UI web.
Alurnya sama persis: BY LINK / BY PROFILE / MIXED / keyword, filter, komentar,
enrichment profil, impor, sentimen, lalu simpan Excel/CSV ke folder outputs/.
"""
import os

from . import core
from .core import *  # noqa: F401,F403  (fungsi & konstanta scraping)
from .core import (_REQN, _CKPT, _checkpoint_path, _text_matches_terms, _lang_ok,
                   _text_has_negative, _json)


ALL_COLS = ["platform", "type", "parent_id", "date", "time", "username", "user_id", "profile_image", "text", "likes", "comments", "shares", "views", "followers", "id", "url", "query"]


class _Label:
    def __init__(self):
        self.value = "starting..."


class Progress:
    """Pengganti widgets.IntProgress + Label: dibaca UI lewat polling."""
    def __init__(self):
        self.min = 0
        self.max = 1
        self.value = 0
        self.bar_style = "info"
        self.label = _Label()

    def to_dict(self):
        return {"value": self.value, "max": self.max, "bar_style": self.bar_style, "label": self.label.value}


DEFAULTS = {
    "mode": "ALL platforms", "queries": "", "limit": 30, "budget": 200,
    "time_filter": False, "since_date": "", "since_time": "", "until_date": "", "until_time": "",
    "strict": True, "lang": "", "neg_use": False, "neg_terms": "",
    "dedup": True, "coverage": False, "delta": False,
    "comments": False, "cmt_top": 20, "cmt_max": 30, "cmt_kw": False, "cmt_kw_txt": "",
    "followers": False, "fill_id": True,
    "tw_rt": True, "tw_rt_cap": 0, "tw_rt_total": 0, "tw_rt_link": False,
    "nlp": "ai", "import_text": "", "raw": False,
    "columns": None, "fmt": "Excel (.xlsx)",
}



def parse_date(s):
    try:
        return datetime.strptime((s or "").strip(), "%Y-%m-%d").date()
    except Exception:
        return None


def wib_epoch(dstr, tstr, is_end):
    d = parse_date(dstr)
    if not d:
        return None
    ts = (tstr or "").strip()
    if ts:
        hh, mm = (ts.split(":") + ["0"])[:2]
        dt = datetime(d.year, d.month, d.day, int(hh), int(mm), 0, tzinfo=WIB)
    else:
        dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=WIB)
        if is_end:
            dt += timedelta(days=1)
    return int(dt.timestamp())


def run_job(params, progress, result):
    """Jalankan satu scraping. `result` diisi files/rows/preview saat selesai."""
    P = dict(DEFAULTS)
    P.update({k: v for k, v in (params or {}).items() if k in DEFAULTS})
    for k in ("limit", "budget", "cmt_top", "cmt_max", "tw_rt_cap", "tw_rt_total"):
        try:
            P[k] = int(P[k] or 0)
        except (TypeError, ValueError):
            P[k] = DEFAULTS[k]
    if not P["columns"]:
        P["columns"] = list(ALL_COLS)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if not core.API_KEY:
        print("API key is empty. Isi SCRAPECREATORS_API_KEY di file .env atau di kolom API key pada UI."); return

    _REQN["total"] = 0; _REQN["fase"] = {}; _REQN["twitter"] = 0; _REQN["other"] = 0
    TW_INC_RT["on"] = bool(P["tw_rt"])
    TW_INC_RT["cap_per_orig"] = max(0, int(P["tw_rt_cap"] or 0))
    TW_INC_RT["cap_total"] = max(0, int(P["tw_rt_total"] or 0))
    NLP_MODE["mode"] = P["nlp"]
    STRICT_KW["on"] = bool(P["strict"])
    STRICT_KW["terms"] = []
    TW_LANG["code"] = P["lang"]
    LANG_FILTER["mode"] = P["lang"]
    NEG_KW["terms"] = ([w.strip() for w in (P["neg_terms"] or "").replace("\n", ",").split(",") if w.strip()]
                       if P["neg_use"] else [])
    BUDGET["max"] = max(0, int(P["budget"] or 0)); BUDGET["warned"] = False
    CURPHASE["name"] = "cek saldo"
    bal = check_credits()
    def _creditnum(b):
        if isinstance(b, dict):
            for k in ("creditCount", "credits", "credit_balance", "balance"):
                v = b.get(k)
                if isinstance(v, (int, float)):
                    return int(v)
        return None
    start_credits = _creditnum(bal)
    if start_credits is not None:
        print(f"\U0001F4B0 Credit balance: {start_credits:,}")
    else:
        print("Credits:", bal)
    prog = progress
    prog_lbl = progress.label
    print("To stop early: press the ■ Stop button. Data collected so far is still saved.")
    ensure_drive()
    print(f"\U0001F4BE Autosave berkala AKTIF: progres disimpan tiap ~50 baris ke {_checkpoint_path()}")
    try:
        import os as _os
        _cp = _checkpoint_path()
        if _os.path.exists(_cp) and _os.path.getsize(_cp) > 50:
            import pandas as _pd
            _prev_n = len(_pd.read_csv(_cp))
            print(f"  \u2139\uFE0F Ditemukan checkpoint run sebelumnya ({_prev_n} baris) di {_cp}.")
            print("     Jika run lalu terputus, data itu masih ada di file tsb. Impor via kotak 'Import file(s)' bila perlu.")
    except Exception:
        pass
    _CKPT["last"] = 0

    platform = P["mode"]
    platforms = ALL_PLATFORMS if platform == "ALL platforms" else [platform]
    mixed = (platform == MIXED_MODE)
    mixed_post_urls, mixed_profile_urls = [], []
    if mixed:
        kw_lines = []
        for l in (x.strip() for x in P["queries"].splitlines()):
            if not l:
                continue
            if l.lower().startswith("http"):
                for u in re.split(r"\s+", l):
                    if u.lower().startswith("http"):
                        (mixed_post_urls if looks_like_post_url(u) else mixed_profile_urls).append(u)
            else:
                kw_lines.append(l)
        mixed_post_urls = list(dict.fromkeys(mixed_post_urls))
        mixed_profile_urls = list(dict.fromkeys(mixed_profile_urls))
        terms = [t.strip() for l in kw_lines for t in l.split(",") if t.strip()]
        STRICT_KW["terms"] = terms
        platforms = ALL_PLATFORMS if terms else []
        print(f"\U0001F500 Mode CAMPUR terdeteksi: {len(terms)} keyword (semua platform), "
              f"{len(mixed_post_urls)} link postingan, {len(mixed_profile_urls)} link profil.")
        if not (terms or mixed_post_urls or mixed_profile_urls):
            print("Masukkan minimal satu keyword atau URL."); return
    else:
        terms = [t.strip() for t in re.split(r"[\n,]+", P["queries"]) if t.strip()]
        STRICT_KW["terms"] = terms
        if not terms:
            print("Enter at least one keyword/hashtag."); return

    since_str = (P["since_date"] or "").strip()
    until_str = (P["until_date"] or "").strip()
    since_tstr = (P["since_time"] or "").strip()
    until_tstr = (P["until_time"] or "").strip()
    since_str = since_str if parse_date(since_str) else ""
    until_str = until_str if parse_date(until_str) else ""
    use_filter = P["time_filter"]
    s_ep = wib_epoch(since_str, since_tstr, False) if use_filter else None
    u_ep = wib_epoch(until_str, until_tstr, True) if use_filter else None

    if use_filter:
        if s_ep is None and u_ep is None:
            print("\u274C Time filter DICENTANG tapi tanggal From/To belum terisi lengkap (Y, M, dan D wajib diisi).")
            print("   Run DIBATALKAN supaya tidak boros kredit. Isi tanggalnya lalu klik Scrape lagi.")
            return
        _fr = f"{since_str} {since_tstr or '00:00'}" if s_ep is not None else "(tanpa batas awal)"
        _to = f"{until_str} {until_tstr or '00:00'}" if u_ep is not None else "(tanpa batas akhir)"
        print(f"\u23F1\uFE0F Filter waktu AKTIF: {_fr} \u2192 {_to} WIB")
        if s_ep is None or u_ep is None:
            print("   \u26A0\uFE0F Salah satu batas kosong/tidak valid \u2014 pastikan memang disengaja.")

    def _in_tf(r):
        """True kalau baris masuk rentang waktu (atau filter mati / benar-benar tanpa tanggal)."""
        if not (use_filter and (s_ep or u_ep)):
            return True
        ep = row_epoch_wib(r.get("date"), r.get("time"))
        if ep is None:
            return True
        if s_ep and ep < s_ep: return False
        if u_ep and ep >= u_ep: return False
        return True

    if use_filter and s_ep and u_ep:
        _range_days = max(1, round((u_ep - s_ep) / 86400))
        if _range_days <= 3 and int(P["limit"]) > 40:
            print(f"\U0001F4A1 Rentang filter cuma {_range_days} hari tapi Max/query = {P["limit"]}.")
            print("   Platform selain Twitter/Threads TIDAK bisa difilter di server \u2014 semua diambil dulu baru disaring (bayar penuh).")
            print("   Saran: kecilkan Max/query jadi 20\u201330 untuk rentang pendek. Batas kredit/run tetap jadi rem pengaman.")

    if platform == LINK_MODE:
        urls = [u.strip() for u in re.split(r"\s+", P["queries"]) if u.strip().lower().startswith("http")]
        if not urls:
            if P["queries"].strip():
                print("BY LINK mode only reads URLs (lines starting with http), but you typed keywords.")
                print("To search KEYWORDS: switch Platform to 'ALL platforms' or a specific platform (TikTok/YouTube/etc.), then Scrape again.")
            else:
                print("Paste at least one valid URL.")
            return
        nuniq = len(set(urls))
        mode_txt = "merged" if P["dedup"] else "one row per input"
        print(f"{len(urls)} URL lines, {nuniq} unique links (duplicates {mode_txt}).")
        print(f"\U0001F4B3 Estimated cost: ~{nuniq} credits (1 per unique link). Actual shown at the end.")
        _stl = {"n": 0}
        def _plink(done, total, res):
            prog.max = total; prog.value = done
            if res: _stl["n"] += 1
            prog_lbl.value = f"{done}/{total} links fetched • {_stl['n']} with data"
        CURPHASE["name"] = "by link"
        allrows = scrape_links(urls, P["raw"], P["dedup"], lambda m: print(m), progress=_plink)
        if P["tw_rt_link"]:
            CURPHASE["name"] = "retweeters"
            tw_links = [u for u in urls if "twitter.com" in u or "x.com" in u]
            _rtc = TW_INC_RT.get("cap_per_orig", 0) or 100
            print(f"\n\U0001F4B3 Retweeters: {len(tw_links)} tweet \u00d7 s/d {_rtc} akun (~{len(tw_links)}\u2013{len(tw_links)*(_rtc//25+1)} kredit).")
            for tu in tw_links:
                try:
                    allrows.extend(fetch_retweeters(tu, P["raw"], lambda m: print(m), cap=_rtc))
                except BudgetExceeded:
                    break
        platforms = []
    elif platform == PROFILE_MODE:
        purls = [u.strip() for u in re.split(r"\s+", P["queries"]) if u.strip().lower().startswith("http")]
        purls = list(dict.fromkeys(purls))
        if not purls:
            print("Paste at least one profile URL (e.g. https://www.instagram.com/username/)."); return
        _mq = int(P["limit"])
        pages_est = max(1, -(-_mq // 20))   # ceil(mq/20): 5 atau 20 -> 1 halaman; 40 -> 2; 100 -> 5
        est_lo = len(purls) * pages_est
        print(f"{len(purls)} profile(s).")
        print(f"\U0001F4B3 Estimasi: ~{est_lo} kredit = {len(purls)} profil \u00d7 {pages_est} halaman "
              f"(Max/query={_mq} berlaku PER PROFIL; \u226420 = 1 halaman = ~1 kredit/profil).")
        if _mq <= 20:
            print(f"  \u2705 Hemat: Max/query={_mq} \u2192 ~1 kredit per profil, total \u2248 {len(purls)} kredit.")
        elif _mq > 30 and len(purls) >= 10:
            print(f"  \u26A0\uFE0F Max/query={_mq} \u00d7 {len(purls)} profil = boros. Turunkan ke 5\u201320 untuk hemat. "
                  f"Filter waktu juga menghentikan pagination di luar rentang.")
        if BUDGET["max"]:
            _avail = BUDGET["max"] - _REQN["other"]
            if est_lo > _avail:
                _fit = max(1, _avail // pages_est)
                print(f"  \U0001F6D1 Estimasi ({est_lo}) > sisa anggaran ({_avail}). Dibatasi ke {_fit} profil pertama "
                      f"agar tidak jebol. Naikkan 'Credit cap/run' bila ingin semua {len(purls)}.")
                purls = purls[:_fit]
        prog.max = max(1, len(purls)); prog.value = 0
        _stp = {"rows": 0}
        def _pp(done, total, res):
            prog.max = total; prog.value = done
            if res: _stp["rows"] += len(res)
            prog_lbl.value = f"{done}/{total} profiles • {_stp['rows']} posts collected"
        CURPHASE["name"] = "by profile"
        def _prof_one(u):
            return scrape_profile(u, int(P["limit"]), P["raw"], lambda m: print(m),
                                  since_d=since_str if use_filter else "",
                                  until_d=until_str if use_filter else "",
                                  s_ep=s_ep, u_ep=u_ep)
        results = []
        for _r in parallel_map(_prof_one, purls, workers=4, progress=_pp):
            results.append(_r)
            if _r:
                save_checkpoint([x for lst in results if lst for x in lst])
        allrows = []
        for lst in results:
            allrows.extend(lst or [])
        platforms = []
    else:
        allrows = []
        if mixed and mixed_post_urls:
            print(f"\n--- Fase 1/3: BY LINK ({len(mixed_post_urls)} link, ~{len(mixed_post_urls)} kredit) ---")
            _stl = {"n": 0}
            def _plink(done, total, res):
                prog.max = total; prog.value = done
                if res: _stl["n"] += 1
                prog_lbl.value = f"[link] {done}/{total} • {_stl['n']} with data"
            CURPHASE["name"] = "by link"
            allrows.extend(scrape_links(mixed_post_urls, P["raw"], P["dedup"], lambda m: print(m), progress=_plink))
            if P["tw_rt_link"]:
                CURPHASE["name"] = "retweeters"
                tw_links = [u for u in mixed_post_urls if "twitter.com" in u or "x.com" in u]
                _rtc = TW_INC_RT.get("cap_per_orig", 0) or 100
                print(f"\n\U0001F4B3 Retweeters: {len(tw_links)} tweet \u00d7 s/d {_rtc} akun (~{len(tw_links)}\u2013{len(tw_links)*(_rtc//25+1)} kredit).")
                for tu in tw_links:
                    try:
                        allrows.extend(fetch_retweeters(tu, P["raw"], lambda m: print(m), cap=_rtc))
                    except BudgetExceeded:
                        break
        if mixed and mixed_profile_urls:
            pages_est = max(1, int(P["limit"]) // 20)
            if BUDGET["max"]:
                _avail = BUDGET["max"] - _REQN["other"]
                _need = len(mixed_profile_urls) * pages_est
                if _need > _avail:
                    _fit = max(1, _avail // pages_est)
                    print(f"  \U0001F6D1 Profil dibatasi ke {_fit} (sisa anggaran {_avail}).")
                    mixed_profile_urls = mixed_profile_urls[:_fit]
            print(f"\n--- Fase 2/3: BY PROFILE ({len(mixed_profile_urls)} profil, ~{len(mixed_profile_urls) * pages_est} kredit) ---")
            prog.max = max(1, len(mixed_profile_urls)); prog.value = 0
            _stp = {"rows": 0}
            def _pp(done, total, res):
                prog.max = total; prog.value = done
                if res: _stp["rows"] += len(res)
                prog_lbl.value = f"[profil] {done}/{total} • {_stp['rows']} posts"
            CURPHASE["name"] = "by profile"
            results = parallel_map(lambda u: scrape_profile(u, int(P["limit"]), P["raw"], lambda m: print(m),
                                                            since_d=since_str if use_filter else "",
                                                            until_d=until_str if use_filter else "",
                                                            s_ep=s_ep, u_ep=u_ep),
                                   mixed_profile_urls, workers=4, progress=_pp)
            for lst in results:
                allrows.extend(lst or [])
        if mixed and terms:
            print(f"\n--- Fase 3/3: BY KEYWORD ({len(terms)} keyword \u00d7 semua platform) ---")
    seen = set()
    def _expand_tasks():
        out = []
        for plat in platforms:
            for term in terms:
                t = term
                if "hashtag" in plat.lower():
                    t = term.replace(" ", "")
                out.append((plat, t))
        if P["coverage"] and platform in ("ALL platforms", MIXED_MODE):
            for term in terms:
                nospace = term.replace(" ", "")
                extra = [("TikTok (hashtag)", nospace), (IG_KW_PLATFORM, term)]
                if core.TWITTER_API_KEY.strip():
                    extra.append((TW_PLATFORM, term))
                for e in extra:
                    if e not in out:
                        out.append(e)
        return list(dict.fromkeys(out))
    COVERAGE["on"] = bool(P["coverage"])
    tasks = _expand_tasks()
    if P["coverage"]:
        print(f"\U0001F50D Max coverage ON \u2014 {len(tasks)} search tasks (keyword + hashtag + Google + Twitter Latest&Top)")
    _pages_est = max(1, int(P["limit"]) // 20)
    _google_tasks = sum(1 for t in tasks if "Google" in t[0])
    _api_tasks = len(tasks) - _google_tasks
    _lo = _api_tasks * 1 + _google_tasks * 2
    _hi = _api_tasks * _pages_est + _google_tasks * ((4 if P["coverage"] else 1) * 6 + min(int(P["limit"]), 40))
    _cmt = ""
    if P["comments"]:
        _tn = int(P["cmt_top"] or 20)
        _cmt = f" + comments ~{_tn if _tn > 0 else 'all posts \u00d7 pages'}"
    print(f"\U0001F4B3 Estimasi: ~{_lo}\u2013{_hi} kredit{_cmt} ({len(tasks)} tugas cari). Angka pasti di akhir run.")
    if BUDGET["max"] and _lo > (BUDGET["max"] - _REQN["other"]):
        print(f"  \U0001F6D1 Estimasi minimum ({_lo}) sudah melebihi sisa anggaran ({BUDGET['max'] - _REQN['total']}).")
        print(f"     Run akan berhenti otomatis di batas {BUDGET['max']} \u2014 sebagian keyword mungkin tak terproses.")
        print(f"     Saran: kurangi keyword, atau naikkan 'Credit cap/run'.")
    elif _hi > 800:
        print(f"  \u26A0\uFE0F Estimasi maksimum tinggi (~{_hi}). Untuk hemat: kecilkan Max/query, kurangi keyword,")
        print(f"     matikan 'Max coverage', dan pastikan filter waktu (Last 24 hours) aktif.")
    if tasks:
        prog.max = max(1, len(tasks)); prog.value = 0
    _stsearch = {"rows": 0}
    def _psearch(done, total, res):
        prog.max = total; prog.value = done
        if res and len(res) >= 3:
            _stsearch["rows"] += len(res[2])
        prog_lbl.value = f"{done}/{total} tasks done • {_stsearch['rows']} rows collected"

    def _run_task(pt):
        plat, term = pt
        buf = []
        def _go():
            return scrape(plat, term, int(P["limit"]),
                          since_str if P["time_filter"] else "",
                          until_str if P["time_filter"] else "",
                          P["raw"], lambda m: buf.append(m),
                          s_ep=s_ep, u_ep=u_ep)
        try:
            rows = _go()
            if not rows:
                buf.append("  0 results \u2192 auto-retry once...")
                time.sleep(3)
                rows = _go()
        except Exception as e:
            rows = []; buf.append(f"  (error: {e})")
        return (plat, term, rows, buf)

    _filtered_out = {"n": 0}
    _strict_out = {"n": 0}
    _lang_out = {"n": 0}
    _neg_out = {"n": 0}
    CURPHASE["name"] = "keyword search"
    for plat, term, rows, buf in parallel_map(_run_task, tasks, workers=4, progress=_psearch):
        print(f"\n=== {plat} :: {term} ===")
        for line in buf:
            print(line)
        for r in rows:
            key = (r["platform"], r["id"] or r["text"][:50])
            if key in seen:
                continue
            if STRICT_KW["on"] and r.get("type") != "comment" and \
               not _text_matches_terms(r.get("text", ""), STRICT_KW["terms"]):
                _strict_out["n"] += 1
                continue
            if not _lang_ok(r.get("text", "")):
                _lang_out["n"] += 1
                continue
            if r.get("type") != "comment" and _text_has_negative(r.get("text", ""), NEG_KW["terms"]):
                _neg_out["n"] += 1
                continue
            if use_filter and (s_ep or u_ep):
                _ep = row_epoch_wib(r.get("date"), r.get("time"))
                if _ep is not None:
                    if s_ep and _ep < s_ep: _filtered_out["n"] += 1; continue
                    if u_ep and _ep >= u_ep: _filtered_out["n"] += 1; continue
            seen.add(key)
            allrows.append(r)
        save_checkpoint(allrows)
    save_checkpoint(allrows, force=True)
    if use_filter and _filtered_out["n"]:
        print(f"\n\u26A0\uFE0F Time filter removed {_filtered_out['n']} rows. If clips seem missing, widen or disable the time filter.")
    if STRICT_KW["on"] and _strict_out["n"]:
        print(f"\U0001F3AF Strict keyword: {_strict_out['n']} baris dibuang (tak memuat frasa kata kunci utuh).")
    if LANG_FILTER["mode"] and _lang_out["n"]:
        _lname = {"id": "Indonesia", "en": "Inggris"}.get(LANG_FILTER["mode"], LANG_FILTER["mode"])
        print(f"\U0001F310 Filter bahasa: {_lang_out['n']} baris non-{_lname} dibuang (Thai/Arab/Mandarin/Turki/dll).")
    if NEG_KW["terms"] and _neg_out["n"]:
        print(f"\U0001F6AB Negative keyword: {_neg_out['n']} baris noise dibuang (K-pop/motor/jualan/game/dll).")

    for r in allrows:
        r.setdefault("type", "post"); r.setdefault("parent_id", "")
    if use_filter and (s_ep or u_ep):
        _b = len(allrows)
        allrows = [r for r in allrows if _in_tf(r)]
        if _b - len(allrows):
            print(f"\u23F1\uFE0F Time filter (semua mode): {_b - len(allrows)} baris di luar rentang dibuang.")
    _seenk, _uniq = set(), []
    for r in allrows:
        k = (r.get("platform"), r.get("type"),
             str(r.get("id") or r.get("url") or (r.get("text") or "")[:80]))
        if k in _seenk:
            continue
        _seenk.add(k); _uniq.append(r)
    if len(_uniq) < len(allrows):
        print(f"\U0001F9F9 Dedup global: {len(allrows) - len(_uniq)} baris duplikat dibuang.")
    allrows = _uniq
    if P["delta"] and allrows:
        prev = load_seen_ids()
        keyed = [(f"{r['platform']}:{r.get('id') or r.get('text','')[:60]}", r) for r in allrows]
        fresh = [r for k, r in keyed if k not in prev]
        dropped = len(allrows) - len(fresh)
        if dropped:
            print(f"\U0001F195 Delta mode: {dropped} rows already seen in previous runs \u2192 keeping {len(fresh)} NEW rows.")
        allrows = fresh
        save_seen_ids(prev | {k for k, _ in keyed})
    if P["comments"] and allrows:
        post_rows = [r for r in allrows if r.get("type") != "comment" and r.get("url")
                     and (COMMENTS_EP.get(r["platform"])
                          or r["platform"] == "Threads"
                          or (r["platform"] == "Twitter" and core.TWITTER_API_KEY.strip() and r.get("id")))]
        topn = int(P["cmt_top"] if P["cmt_top"] is not None else 20)
        if topn > 0 and len(post_rows) > topn:
            post_rows = sorted(post_rows, key=lambda r: (r.get("likes") or 0) + (r.get("comments") or 0), reverse=True)[:topn]
        capn = int(P["cmt_max"] if P["cmt_max"] is not None else 30)
        # --- Estimasi biaya komentar + rem keras ---
        pages_per_post = 1 if capn <= 0 else max(1, (capn + 24) // 25)  # ~25 komentar/halaman
        if capn <= 0:
            pages_per_post = 8   # asumsi konservatif utk "semua komentar"
        est_credits = len(post_rows) * pages_per_post
        cap_txt = "ALL comments (mahal!)" if capn <= 0 else f"max {capn} comments/post"
        print(f"\n\U0001F4B3 Estimasi komentar: {len(post_rows)} post \u00d7 ~{pages_per_post} req = ~{est_credits} kredit ({cap_txt}).")
        _cap_left = (BUDGET["max"] - _REQN["other"]) if BUDGET["max"] else 10**9
        if BUDGET["max"] and est_credits > _cap_left:
            # potong jumlah post agar muat dalam sisa anggaran
            _fit = max(0, _cap_left // max(1, pages_per_post))
            if _fit < len(post_rows):
                print(f"  \u26A0\uFE0F Sisa anggaran kredit cuma {_cap_left} \u2014 komentar dibatasi ke {_fit} post teratas (dari {len(post_rows)}) agar tidak jebol.")
                post_rows = post_rows[:_fit]
                est_credits = len(post_rows) * pages_per_post
        if capn <= 0 and len(post_rows) > 5:
            print(f"  \u26A0\uFE0F 'Max cmts/post = 0' (semua komentar) sangat boros untuk {len(post_rows)} post. Sangat disarankan isi 30\u2013100.")
        print(f"Fetching comments for {len(post_rows)} posts...")
        _tw_n = sum(1 for r in post_rows if r["platform"] == "Twitter")
        if _tw_n:
            print(f"  (termasuk {_tw_n} tweet \u2014 reply diambil via pencarian conversation_id, \u00b11+ kredit Twitter/tweet)")
        CURPHASE["name"] = "komentar/reply"
        clists = parallel_map(lambda r: (fetch_tw_replies(r.get("id"), P["raw"], lambda m: None, cap=capn)
                                         if r["platform"] == "Twitter"
                                         else fetch_comments(r["platform"], r["url"], r.get("id"), P["raw"], lambda m: None, cap=capn)),
                              post_rows, workers=6)
        kwfilter = []
        if P["cmt_kw"]:
            kwfilter = [w.strip().lower() for w in (P["cmt_kw_txt"] or "").split(",") if w.strip()]
            if not kwfilter:
                kwfilter = [t.lower() for t in terms if t]
            if not kwfilter:
                print("  \u26A0\uFE0F Filter komentar aktif tapi tidak ada kata kunci (mode link/profil tanpa kata filter) \u2014 filter dilewati.")
        added, dropped_kw = 0, 0
        for cl in clists:
            if not cl:
                continue
            if kwfilter:
                keep = [c for c in cl if any(k in str(c.get("text", "")).lower() for k in kwfilter)]
                dropped_kw += len(cl) - len(keep)
                cl = keep
            cl = [c for c in cl if _in_tf(c)]
            allrows.extend(cl); added += len(cl)
        if kwfilter:
            print(f"  +{added} komentar mengandung kata kunci ({dropped_kw} komentar lain dibuang)")
        else:
            print(f"  +{added} comment rows")
    if (P["followers"] or P["fill_id"]) and allrows:
        need = set()
        for r in allrows:
            ident = r.get("_profile") or r.get("username")
            if not ident or not PROFILE_EP.get(r["platform"]):
                continue
            need_foll = P["followers"] and r.get("type") != "comment" and not r.get("followers")
            need_id = P["fill_id"] and (not r.get("user_id") or not r.get("profile_image"))
            if need_foll or need_id:
                need.add((r["platform"], ident))
        need = list(need)
        if need:
            what = " + ".join([w for w, on in (("followers", P["followers"]),
                                               ("user id/image", P["fill_id"])) if on])
            print(f"\n\U0001F50E Melengkapi profil ({what}) untuk {len(need)} akun unik \u2014 1 request/akun, digabung satu fase...")
            CURPHASE["name"] = "profil enrichment"
            prog.max = max(1, len(need)); prog.value = 0
            def _pprof(done, total, res):
                prog.max = total; prog.value = done
                prog_lbl.value = f"[profil] {done}/{total} akun"
            vals = parallel_map(lambda pu: (pu, fetch_profile(pu[0], pu[1], lambda m: None)),
                                need, workers=12, progress=_pprof)
            pmap = {pu: v for pu, v in vals if pu is not None and v}
            for r in allrows:
                ident = r.get("_profile") or r.get("username")
                p = pmap.get((r["platform"], ident))
                if not p:
                    continue
                if P["followers"] and r.get("type") != "comment" and not r.get("followers") \
                   and p.get("followers") is not None:
                    r["followers"] = p["followers"]
                if P["fill_id"]:
                    if not r.get("user_id") and p.get("user_id"):
                        r["user_id"] = p["user_id"]
                    if not r.get("profile_image") and p.get("profile_image"):
                        r["profile_image"] = p["profile_image"]
    for r in allrows:
        if not r.get("user_id") and r.get("username"):
            r["user_id"] = r["username"]  # fallback terakhir: pakai username agar kolom tidak kosong
    # User ID Parent: user id pemilik unggahan induk (untuk cluster komentar per unggahan)
    _pid_owner = {str(r.get("id")): r.get("user_id", "")
                  for r in allrows if r.get("type") != "comment" and r.get("id")}
    for r in allrows:
        if r.get("type") == "comment":
            r["user_id_parent"] = _pid_owner.get(str(r.get("parent_id") or ""), "")
        else:
            r["user_id_parent"] = r.get("user_id", "")

    if not allrows:
        print("\nNo results. Check the '(response: ...)' lines above.")
        print("Tip: turn off the time filter, check your credits, or change platform.")
        return

    # ---- Impor data yang sudah dimiliki (tanpa crawling, 0 kredit) ----
    _imp_lines = [p.strip() for p in (P["import_text"] or "").splitlines() if p.strip()]
    if _imp_lines:
        _imp_rows = []
        _colmap = {"username":"username","user":"username","user id":"user_id","userid":"user_id",
                   "text":"text","isi":"text","caption":"text","date":"date","tanggal":"date",
                   "time":"time","jam":"time","likes":"likes","like":"likes","reply":"comments",
                   "comment":"comments","comments":"comments","shares":"shares","share":"shares",
                   "url":"url","link":"url","platform":"platform","sub type":"type","type":"type",
                   "views count":"views","views":"views","follower":"followers","followers":"followers",
                   "parent id":"parent_id","id":"id","user image":"profile_image"}
        _keep = ("username","user_id","text","date","time","likes","comments","shares",
                 "url","platform","type","views","followers","parent_id","id","profile_image")

        def _plat_from_url(u):
            u = str(u).lower()
            if "tiktok.com" in u: return "tiktok"
            if "youtube.com" in u or "youtu.be" in u: return "youtube"
            if "instagram.com" in u: return "instagram"
            if "facebook.com" in u or "fb.watch" in u: return "facebook"
            if "threads." in u: return "threads"
            if "twitter.com" in u or "x.com" in u: return "twitter"
            return "imported"

        # Pisahkan: baris path file (.xlsx/.csv) vs baris data mentah (link / link<TAB/koma>teks)
        _paths = [l for l in _imp_lines if l.lower().endswith((".xlsx", ".xls", ".csv"))]
        _raw   = [l for l in _imp_lines if l not in _paths]

        # 1) File yang sudah dimiliki
        for pth in _paths:
            try:
                idf = pd.read_csv(pth) if pth.lower().endswith(".csv") else pd.read_excel(pth)
                idf.columns = [str(c).strip() for c in idf.columns]
                idf = idf.rename(columns={c: _colmap[c.lower()] for c in idf.columns if c.lower() in _colmap})
                for _, rr in idf.iterrows():
                    d = {k: rr[k] for k in idf.columns if k in _keep}
                    d.setdefault("platform", "imported"); d.setdefault("type", "post")
                    if d.get("url") and (not d.get("platform") or d["platform"] in ("imported","")):
                        d["platform"] = _plat_from_url(d["url"])
                    d["query"] = "(imported)"; d["_imported"] = True
                    _imp_rows.append(d)
                print(f"  \U0001F4E5 Impor {len(idf)} baris dari file {pth} (0 kredit)")
            except Exception as e:
                print(f"  \u26A0\uFE0F Gagal impor file {pth}: {type(e).__name__}: {str(e)[:60]}")

        # 2) Data mentah yang ditempel langsung: bulk link, opsional diikuti teks
        #    Format tiap baris: <url>   atau   <url><TAB atau ; atau ,><teks bebas>
        for ln in _raw:
            parts = re.split(r"\t|\s*\|\s*|\s*;\s*", ln, maxsplit=1)
            url = parts[0].strip()
            txt = parts[1].strip() if len(parts) > 1 else ""
            if not url.lower().startswith("http"):
                # bukan link & bukan path -> anggap seluruh baris sbg teks tanpa url
                url, txt = "", ln
            d = {"url": url, "text": txt, "platform": _plat_from_url(url) if url else "imported",
                 "type": "post", "query": "(imported)", "_imported": True}
            _imp_rows.append(d)
        if _raw:
            print(f"  \U0001F4CE {len(_raw)} baris data tempel langsung (bulk link) diterima \u2014 TANPA crawling, 0 kredit")

        if _imp_rows:
            allrows.extend(_imp_rows)
            print(f"  \u2705 {len(_imp_rows)} baris impor digabung untuk dianalisis & diformat.")
    if not allrows:
        prog.bar_style = "warning"; print("\nNo results and no imported rows."); return
    prog.bar_style = "success"; prog.value = prog.max
    prog_lbl.value = f"Done • {len(allrows)} rows collected"
    df = pd.DataFrame(allrows)
    df = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")
    chosen = [c for c in ALL_COLS if (c in P["columns"]) and c in df.columns]
    for _core in ("type", "parent_id", "text", "platform"):
        if _core in df.columns and _core not in chosen:
            chosen.append(_core)
    if chosen:
        extras = [c for c in df.columns if c not in ALL_COLS]
        df = df[[c for c in ALL_COLS if c in chosen] + extras]
    # Sentiment (NLP) + Emotion (Plutchik) + layout mengikuti Sample_Crawling
    print(f"\n\U0001F9E0 Analisis Sentimen & Emosi: {'model AI (hybrid + leksikon)' if NLP_MODE['mode'] == 'ai' else 'leksikon'} \u2014 lokal, tanpa kredit API...")
    df = to_sample_layout(df)
    if until_str and P["time_filter"]:
        label = until_str + " " + (until_tstr or "23.59")
    else:
        label = datetime.now(WIB).strftime("%Y-%m-%d %H.%M")
    label = label.replace(":", ".")
    if platform == LINK_MODE:
        plat_label = "ByLink"
    elif platform == MIXED_MODE:
        plat_label = "Mixed"
    elif platform == "ALL platforms":
        plat_label = "All"
    else:
        plat_label = platform.split(" ")[0]
    base = f"Data {plat_label}_{label}"

    saved = []
    if "Excel" in P["fmt"]:
        xp = os.path.join(OUTPUT_DIR, base + ".xlsx")
        df.to_excel(xp, index=False)
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import Font, PatternFill
            wb = load_workbook(xp); ws = wb.active
            hdr_font = Font(bold=True, color="FFFFFF")
            hdr_fill = PatternFill("solid", fgColor="4472C4")
            for c in ws[1]:
                c.font = hdr_font; c.fill = hdr_fill
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            widths = {"No": 6, "Sentiment": 11, "Username": 18, "User ID": 16, "Follower": 11,
                      "Text": 60, "Date": 11, "Time": 9, "Likes": 9, "Reply": 9, "Shares": 9,
                      "Url": 40, "Emotion": 12, "User Image": 30, "Platform": 10, "Sub Type": 10,
                      "Views Count": 12, "Gender": 9, "Age Category": 13, "Content Location": 16,
                      "Religion": 9, "Ethnic": 9, "Phsychography": 14, "Parent ID": 20,
                      "ID": 20, "User ID Parent": 18, "Query": 16, "Raw JSON": 30}
            for i, col in enumerate(df.columns, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = widths.get(col, 14)
            wb.save(xp)
        except Exception:
            pass
        try:
            add_rekap_codebook(xp, df)
            print("   \U0001F4D1 Sheet 'Rekap' + 'Codebook' ditambahkan ke Excel.")
        except Exception as e:
            print(f"   (Rekap dilewati: {type(e).__name__})")
        saved.append(xp)
    if "CSV" in P["fmt"]:
        cp = os.path.join(OUTPUT_DIR, base + ".csv")
        df.to_csv(cp, index=False); saved.append(cp)
    print(f"\n\u2705 Done. {len(df)} rows -> {', '.join(saved)}")
    try:
        counts = df["Platform"].value_counts().to_dict()
        print("   Per platform: " + ", ".join(f"{k}: {v}" for k, v in counts.items()))
    except Exception:
        pass
    CURPHASE["name"] = "cek saldo"
    print_usage_breakdown()
    end_credits = _creditnum(check_credits())
    if start_credits is not None and end_credits is not None:
        print(f"   \U0001F4B0 Kredit terpakai run ini: {start_credits - end_credits:,} \u2022 sisa: {end_credits:,}")
    result["files"] = [os.path.basename(f) for f in saved]
    result["rows"] = len(df)
    try:
        result["preview_columns"] = [str(c) for c in df.columns]
        result["preview"] = _json.loads(df.head(30).to_json(orient="values", force_ascii=False, date_format="iso"))
    except Exception:
        pass
    print(f"\U0001F4C1 File tersimpan di folder: {OUTPUT_DIR}")
