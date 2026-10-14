"""Social Scraper — core logic (ported from the Colab notebook v58).

Scraping, parsing, filters, sentiment and the Excel layout are taken from the notebook;
only the Colab-specific parts (widgets, Google Drive, files.download) were replaced so it
runs on a laptop. API keys are read from the .env file / environment.
"""
import os

API_KEY = (os.environ.get("SCRAPECREATORS_API_KEY") or "").strip()
TWITTER_API_KEY = (os.environ.get("TWITTER_API_KEY") or "").strip()
OUTPUT_DIR = os.path.abspath(os.environ.get("OUTPUT_DIR") or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs"))


def set_api_keys(api_key=None, twitter_key=None):
    """Change API keys at runtime (from the UI) without restarting the server."""
    global API_KEY, TWITTER_API_KEY
    if api_key is not None:
        API_KEY = api_key.strip()
        HEADERS["x-api-key"] = API_KEY
    if twitter_key is not None:
        TWITTER_API_KEY = twitter_key.strip()

import subprocess, sys
try:
    import openpyxl
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])

import requests, time, json as _json
from datetime import datetime, timedelta, timezone
import pandas as pd

import base64 as _b64
BASE = _b64.b64decode("aHR0cHM6Ly9hcGkuc2NyYXBlY3JlYXRvcnMuY29t").decode()
HEADERS = {"x-api-key": API_KEY}

# ---- Akuntansi request + rem anggaran kredit ----
_REQN = {"total": 0, "fase": {}, "twitter": 0, "other": 0}
CURPHASE = {"name": "setup"}
BUDGET = {"max": 0, "warned": False}

class BudgetExceeded(Exception):
    pass

class StopRequested(BudgetExceeded):
    """Raised by the Stop button: further requests are cancelled, collected data is kept."""
    pass

STOP = {"flag": False}
NET_RETRIES = 3

# Domain penyedia Twitter (getxapi) — request ke sini TIDAK dihitung ke batas kredit,
# karena pakai saldo/paket terpisah. Twitter selalu "unlimited" terhadap Credit cap.
_TWITTER_PROVIDER_HOST = "getxapi"

# Penghitung request + rem anggaran — SELF-HEALING:
# fungsi asli diambil dari requests.api.get (tidak pernah kita sentuh), lalu requests.get
# ditimpa ulang setiap cell dijalankan. Tumpukan pembungkus lama dari run sebelumnya
# otomatis terbuang — aman dijalankan berulang kali TANPA restart runtime.
from requests import api as _rq_api
def _counting_get(*a, _orig=_rq_api.get, **k):
    url = a[0] if a else k.get("url", "")
    if STOP["flag"]:
        raise StopRequested("stopped by user")
    is_twitter = _TWITTER_PROVIDER_HOST in str(url)
    # Rem anggaran HANYA untuk penyedia non-Twitter (ScrapeCreators).
    if (not is_twitter) and BUDGET["max"] and _REQN["other"] >= BUDGET["max"]:
        if not BUDGET["warned"]:
            BUDGET["warned"] = True
            print(f"\n\U0001F6D1 CREDIT CAP REACHED ({BUDGET['max']} non-Twitter requests) \u2014 non-Twitter scraping stopped; Twitter continues; collected data is kept.")
        raise BudgetExceeded("credit cap per run reached (non-Twitter)")
    _REQN["total"] += 1
    if is_twitter:
        _REQN["twitter"] += 1
    else:
        _REQN["other"] += 1
    ph = CURPHASE["name"]
    _REQN["fase"][ph] = _REQN["fase"].get(ph, 0) + 1
    # Retry transient failures (dropped connection, read timeout, 502/503/504) so one hiccup
    # on home Wi-Fi does not lose a whole task. 429 is left to the callers' own back-off.
    for attempt in range(NET_RETRIES + 1):
        try:
            r = _orig(*a, **k)
            if r.status_code in (502, 503, 504) and attempt < NET_RETRIES:
                time.sleep(3 * (attempt + 1))
                continue
            return r
        except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError):
            if attempt >= NET_RETRIES or STOP["flag"]:
                raise
            time.sleep(3 * (attempt + 1))
        if STOP["flag"]:
            raise StopRequested("stopped by user")
_counting_get._is_counting_wrapper = True
requests.get = _counting_get

def print_usage_breakdown():
    if not _REQN["total"]:
        return
    parts = ", ".join(f"{k}: {v}" for k, v in sorted(_REQN["fase"].items(), key=lambda x: -x[1]))
    print(f"\U0001F4CB API requests this run \u2014 total {_REQN['total']} ({parts})")
    print(f"   \U0001F4B3 Non-Twitter (counted toward cap): {_REQN['other']} \u2022 Twitter/getxapi (separate balance): {_REQN['twitter']}")
    print("   (1 request \u2248 1 credit for most endpoints; exact usage = balance difference below)")
WIB = timezone(timedelta(hours=7))
from concurrent.futures import ThreadPoolExecutor, as_completed

def parallel_map(fn, items, workers=8, progress=None):
    items = list(items)
    if not items:
        return []
    out = [None] * len(items)
    done = 0
    ex = ThreadPoolExecutor(max_workers=min(workers, len(items)))
    futs = {ex.submit(fn, it): i for i, it in enumerate(items)}
    try:
        for f in as_completed(futs):
            i = futs[f]
            try:
                out[i] = f.result()
            except Exception:
                out[i] = None
            done += 1
            if progress:
                try:
                    progress(done, len(items), out[i])
                except Exception:
                    pass
    except KeyboardInterrupt:
        print("\n⏹ Stopped by user — keeping data collected so far...")
    finally:
        try:
            ex.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            ex.shutdown(wait=False)
    return out

PLATFORMS = {
    "TikTok (keyword)":    {"path": "/v1/tiktok/search/keyword",    "param": "query"},
    "TikTok (hashtag)":    {"path": "/v1/tiktok/search/hashtag",    "param": "hashtag"},
    "TikTok (top)":        {"path": "/v1/tiktok/search/top",        "param": "query"},
    "YouTube (keyword)":   {"path": "/v1/youtube/search",           "param": "query"},
    "Threads (keyword)":   {"path": "/v1/threads/search",           "param": "query", "dates": True},
    "Instagram (hashtag)": {"path": "/v1/instagram/search/hashtag", "param": "hashtag"},
}

FB_PLATFORM = "Facebook (Google)"
IG_KW_PLATFORM = "Instagram (Google keyword)"
TW_PLATFORM = "Twitter (keyword)"
ALL_PLATFORMS = ["TikTok (keyword)", "YouTube (keyword)", "Threads (keyword)", "Instagram (hashtag)", FB_PLATFORM, TW_PLATFORM]

PREFERRED_KEYS = ["posts", "videos", "results", "items", "medias", "aweme_list",
                  "search_item_list", "data", "feed", "tweets", "entries"]


# ---------- parser auto-adaptif ----------
import re

WRAPPERS = ["aweme_info", "item", "content", "node", "video_info", "post_info", "post", "media"]

def unwrap(it):
    cur = it
    for _ in range(4):
        if isinstance(cur, dict):
            nxt = None
            for w in WRAPPERS:
                if w in cur and isinstance(cur[w], dict):
                    nxt = cur[w]; break
            if nxt is None:
                break
            cur = nxt
        else:
            break
    return cur

def parse_relative(s):
    m = re.search(r"(\d+)\s*(year|month|week|day|hour|minute|second|tahun|bulan|minggu|hari|jam|menit|detik)", str(s).lower())
    if not m:
        return None
    n = int(m.group(1)); unit = m.group(2)
    days = {"year":365,"tahun":365,"month":30,"bulan":30,"week":7,"minggu":7,"day":1,"hari":1}.get(unit)
    if days:
        return datetime.now(WIB) - timedelta(days=n*days)
    secs = {"hour":3600,"jam":3600,"minute":60,"menit":60,"second":1,"detik":1}.get(unit)
    if secs:
        return datetime.now(WIB) - timedelta(seconds=n*secs)
    return None

def find_items(data):
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        for k in PREFERRED_KEYS:
            v = data.get(k)
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
        best = []
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict) and len(v) > len(best):
                best = v
        if best:
            return best
        for v in data.values():
            if isinstance(v, dict):
                r = find_items(v)
                if len(r) > len(best):
                    best = r
        return best
    return []


def find_num(obj, subs, depth=0):
    if depth > 6:
        return None
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, (int, float)) and not isinstance(v, bool) and any(s in k.lower() for s in subs):
                return v
        for v in obj.values():
            r = find_num(v, subs, depth + 1)
            if r is not None:
                return r
    elif isinstance(obj, list):
        for v in obj[:20]:
            r = find_num(v, subs, depth + 1)
            if r is not None:
                return r
    return None


def g(item, *paths):
    for p in paths:
        cur = item
        ok = True
        for part in p.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                ok = False
                break
        if ok and cur not in (None, "", []):
            return cur
    return None


LIKE_PATHS   = ["legacy.favorite_count","favorite_count","statistics.digg_count","statistics.diggCount","stats.diggCount","statistics.likeCount","like_count","likeCount","likes","favoriteCount","favorite_count","edge_liked_by.count","edge_media_preview_like.count"]
COMMENT_PATHS= ["legacy.reply_count","reply_count","edge_media_to_parent_comment.count","edge_media_preview_comment.count","statistics.comment_count","statistics.commentCount","stats.commentCount","comment_count","commentCount","comments","edge_media_to_comment.count","direct_reply_count"]
SHARE_PATHS  = ["legacy.retweet_count","retweet_count","retweetCount","statistics.share_count","statistics.shareCount","stats.shareCount","share_count","shareCount","reshare_count","repost_count","shares"]
VIEW_PATHS   = ["views.count","video_play_count","video_view_count","statistics.play_count","statistics.playCount","stats.playCount","play_count","playCount","view_count","viewCount","viewCountInt","views","play_count_int","video_view_count"]

def collect_nums(obj, subs, out, depth=0):
    if depth > 6:
        return
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, (int, float)) and not isinstance(v, bool) and any(s in k.lower() for s in subs):
                out.append(v)
            else:
                collect_nums(v, subs, out, depth + 1)
    elif isinstance(obj, list):
        for v in obj[:20]:
            collect_nums(v, subs, out, depth + 1)

def num_at(item, paths, subs):
    for p in paths:
        v = g(item, p)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return int(v)
        if isinstance(v, str) and v.strip().replace(",", "").isdigit():
            return int(v.strip().replace(",", ""))
    out = []
    collect_nums(item, subs, out)
    return int(max(out)) if out else 0

FOLLOWER_PATHS = ["core.user_results.result.legacy.followers_count","owner.edge_followed_by.count","author.follower_count","author.followerCount","author.followers","user.follower_count","user.followerCount","owner.follower_count","channel.subscriberCount","channel.subscriberCountInt","channel.subscribers","subscriberCount","subscriberCountInt","follower_count","followerCount","followers","edge_followed_by.count"]
FOLLOWER_TEXT_PATHS = ["channel.subscriberCountText","author.followerCountText","subscriberCountText","followerCountText"]

def parse_count(s):
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return int(s)
    if not isinstance(s, str):
        return None
    m = re.search(r"([\d.,]+)\s*(k|m|b|jt|juta|rb|ribu)?", s.strip().lower())
    if not m:
        return None
    raw, unit = m.group(1), (m.group(2) or "")
    if unit:
        np_ = raw.replace(",", ".")
        if np_.count(".") > 1:
            i = np_.find("."); np_ = np_[:i+1] + np_[i+1:].replace(".", "")
        try:
            num = float(np_)
        except Exception:
            return None
        mult = {"k":1e3,"rb":1e3,"ribu":1e3,"m":1e6,"jt":1e6,"juta":1e6,"b":1e9}.get(unit, 1)
        return int(num * mult)
    digits = re.sub(r"[^\d]", "", raw)
    return int(digits) if digits else None

def get_followers(item):
    for p in FOLLOWER_PATHS:
        v = g(item, p)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return int(v)
    for p in FOLLOWER_TEXT_PATHS:
        c = parse_count(g(item, p))
        if c:
            return c
    out = []
    collect_nums(item, ["follower", "subscriber"], out)
    return int(max(out)) if out else None

def prof_ident(plat, it, user):
    if str(plat).startswith("Facebook"):
        return g(it, "author.url", "author.id") or user
    return user

PROFILE_EP = {
    "TikTok":   {"path": "/v1/tiktok/profile",  "param": "handle"},
    "YouTube":  {"path": "/v1/youtube/channel", "param": "handle"},
    "Threads":  {"path": "/v1/threads/profile", "param": "handle"},
    "Instagram":{"path": "/v1/instagram/profile","param": "handle"},
    "Facebook": {"path": "/v1/facebook/profile","param": "url"},
}
_prof_cache = {}

def fetch_profile(plat_short, username, log):
    """Ambil info profil: followers, user_id, profile_image (1 request per akun unik, di-cache)."""
    if not username:
        return None
    key = (plat_short, username)
    if key in _prof_cache:
        return _prof_cache[key]
    ep = PROFILE_EP.get(plat_short)
    val = None
    if ep:
        try:
            r = requests.get(BASE + ep["path"], headers=HEADERS,
                             params={ep["param"]: username}, timeout=30)
            if r.status_code == 200:
                data = r.json()
                node = data.get("user") or data.get("data") or data if isinstance(data, dict) else data
                val = {"followers": get_followers(data),
                       "user_id": get_user_id(data) or get_user_id(node) or str(g(node, "id", "pk", "id_str", "rest_id", "channelId", "uid", "user_id") or ""),
                       "profile_image": get_profile_image(data) or get_profile_image(node)}
            else:
                log(f"    profile @{username}: HTTP {r.status_code}")
        except Exception as e:
            log(f"    profile @{username}: {e}")
    _prof_cache[key] = val
    return val

def fetch_followers(plat_short, username, log):
    p = fetch_profile(plat_short, username, log)
    return p.get("followers") if p else None

def get_text(item):
    t = g(item, "caption.text", "caption", "desc", "title", "text", "description", "content", "message", "legacy.full_text", "full_text")
    if not t:
        edges = g(item, "edge_media_to_caption.edges")
        if isinstance(edges, list) and edges and isinstance(edges[0], dict):
            node = edges[0].get("node")
            if isinstance(node, dict):
                t = node.get("text")
    if isinstance(t, dict):
        t = t.get("text") or t.get("plaintext") or ""
    if isinstance(t, list):
        t = " ".join(str(x.get("plaintext", x) if isinstance(x, dict) else x) for x in t)
    return str(t or "")


_HANDLE_URL_RE = re.compile(r"(?:x\.com|twitter\.com|instagram\.com|tiktok\.com|threads\.(?:net|com)|youtube\.com|facebook\.com)/@?([A-Za-z0-9._-]+)", re.I)
_NOT_HANDLE = {"channel", "c", "user", "p", "reel", "reels", "share", "watch", "status", "video",
               "shorts", "posts", "post", "hashtag", "search", "home", "i", "intl", "explore", "web"}

def _handle_from_url(u):
    """Ekstrak handle/username dari URL profil atau URL post."""
    if not isinstance(u, str) or not u:
        return ""
    s = u.strip()
    if s.startswith("/@"):          # canonicalBaseUrl YouTube, mis. "/@PLN_ID"
        s = s[2:].split("/")[0].split("?")[0]
        return s if s else ""
    m = _HANDLE_URL_RE.search(s)
    if m:
        h = m.group(1)
        if h.lower() not in _NOT_HANDLE:
            return h
    return ""

def get_user(item):
    # 1) Prioritas: field yang benar-benar handle/username (tanpa spasi)
    v = g(item, "user.username", "user.unique_id", "user.uniqueId", "user.screen_name",
          "author.unique_id", "author.uniqueId", "author.userName", "author.username",
          "author.screen_name", "author.screenName", "author.handle",
          "channel.handle", "channelHandle", "author.channelHandle",
          "owner.username", "ownerUsername", "username", "uniqueId",
          "screen_name", "screenName", "userName",
          "core.user_results.result.legacy.screen_name", "core.user_results.result.core.screen_name",
          "legacy.screen_name")
    if isinstance(v, str) and v.strip() and " " not in v.strip():
        return v.strip().lstrip("@")
    # 2) Turunkan dari URL profil yang ada di dalam item
    for p in ("author.url", "author.profile_url", "author.profileUrl", "user.url",
              "channelUrl", "channel.url", "author.channelUrl",
              "canonicalBaseUrl", "author.canonicalBaseUrl", "channel.canonicalBaseUrl"):
        h = _handle_from_url(g(item, p))
        if h:
            return h
    # 3) Fallback terakhir: display name (bisa berspasi) — hanya kalau handle tak ditemukan sama sekali
    v2 = g(item, "user.nickname", "author.nickname", "author.name", "from.name",
           "commenter.name", "channel.title", "nickname", "authorName", "name")
    return str((v if isinstance(v, str) else "") or v2 or "").strip()


def get_id(item):
    return str(g(item, "id", "post_id", "aweme_id", "video_id", "videoId", "pk", "code", "shortcode", "rest_id", "legacy.id_str") or "")

def get_user_id(item):
    return str(g(item,
                 "author.id", "author.uid", "author.userId", "author.user_id",
                 "author.id_str", "author.rest_id", "author.pk",
                 "user.pk", "user.id", "user.id_str", "user.pk_id", "user.rest_id",
                 "owner.id", "owner.pk", "owner.id_str",
                 "channel.id", "channel.channelId", "channelId", "channel_id", "channel.externalId",
                 "core.user_results.result.rest_id", "core.user_results.rest_id",
                 "core.user_results.result.id",
                 "legacy.user_id_str", "user_id_str", "tweet.legacy.user_id_str",
                 "from.id", "commenter.id",
                 "user_id", "userId", "authorId") or "")

def tweet_subtype(item, text=""):
    """Klasifikasi tweet: retweet / quoted / tweet."""
    if g(item, "retweeted_status", "retweeted_tweet", "legacy.retweeted_status_result",
         "legacy.retweeted_status_id_str", "retweeted_status_result"):
        return "retweet"
    if str(text or "").startswith("RT @"):
        return "retweet"
    if g(item, "quoted_status", "quoted_tweet", "quoted_status_result",
         "legacy.quoted_status_id_str", "legacy.quoted_status_result") or \
       g(item, "legacy.is_quote_status", "is_quote_status", "isQuote", "is_quote") in (True, "true", "True", 1):
        return "quoted"
    return "tweet"

def media_subtype(platform, it, url=""):
    """Sub Type per platform: IG carousel/video/image/post; FB video/link/post; lainnya post."""
    p = str(platform)
    u = str(url or "").lower()
    if p.startswith("Instagram"):
        tn = str(g(it, "__typename", "typename") or "").lower()
        pt = str(g(it, "product_type") or "").lower()
        mt = g(it, "media_type", "media.media_type")
        if g(it, "carousel_media", "carousel_media_count", "edge_sidecar_to_children") \
           or mt == 8 or "sidecar" in tn or "carousel" in pt:
            return "carousel"
        if mt == 2 or g(it, "is_video") in (True, 1, "true") or "video" in tn \
           or pt in ("clips", "reels", "igtv") or "/reel" in u:
            return "video"
        if mt == 1 or "image" in tn or "photo" in tn:
            return "image"
        return "post"
    if p.startswith("Facebook"):
        if "/videos/" in u or "/watch" in u or "/reel" in u \
           or g(it, "is_video") in (True, 1, "true") \
           or "video" in str(g(it, "attachments.data.0.type", "attachment.type", "type") or "").lower():
            return "video"
        att = str(g(it, "attachments.data.0.type", "attachment.type") or "").lower()
        if att in ("share", "link") or g(it, "external_url", "link_url"):
            return "link"
        return "post"
    return "post"

def _first_url(node):
    if not node:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, dict):
        for k in ("url_list", "urls"):
            v = node.get(k)
            if isinstance(v, list) and v:
                return v[0] if isinstance(v[0], str) else (v[0].get("url", "") if isinstance(v[0], dict) else "")
            if isinstance(v, str):
                return v
        for k in ("url", "uri", "src"):
            if isinstance(node.get(k), str):
                return node[k]
    if isinstance(node, list) and node:
        first = node[-1] if isinstance(node[-1], dict) else node[0]
        if isinstance(first, dict):
            return first.get("url") or first.get("uri") or ""
        if isinstance(first, str):
            return first
    return ""

def _is_bad_avatar(u):
    ul = u.lower()
    if "i.ytimg.com/vi/" in ul:
        return True
    if "/vi/" in ul and "ytimg" in ul:
        return True
    if "ytimg" in ul and ul.rstrip("/").endswith(("default.jpg", "default_live.jpg")):
        return True
    return False

def get_profile_image(item):
    candidates = []
    for path in ("user.profile_pic_url", "user.profile_pic_url_hd", "owner.profile_pic_url",
                 "owner.profile_pic_url_hd", "profile_pic_url", "profile_pic_url_hd",
                 "author.image", "author.avatar_url", "author.profilePicture",
                 "author.profilePicUrl", "author.profile_pic_url", "author.avatarUrl",
                 "author.profile_image_url", "author.profile_image_url_https",
                 "user.avatar", "user.avatarUrl", "user.profileImageUrl", "profile_picture",
                 "legacy.profile_image_url_https", "user.legacy.profile_image_url_https",
                 "core.user_results.result.avatar.image_url",
                 "core.user_results.result.legacy.profile_image_url_https",
                 "channelThumbnail", "channelAvatar", "channel.avatar", "channel.avatarUrl",
                 "channel.thumbnail", "author.channelThumbnail", "profile_image_url", "avatar_url"):
        v = g(item, path)
        if isinstance(v, str) and v:
            candidates.append(v)
    for base in ("author.avatar_larger", "author.avatar_medium", "author.avatar_thumb",
                 "author.avatar", "channel.avatar.thumbnails", "channel.avatar",
                 "channel.thumbnails", "author.thumbnails", "avatar.thumbnails",
                 "owner.hd_profile_pic_url_info", "user.hd_profile_pic_url_info", "thumbnails"):
        u = _first_url(g(item, base))
        if u:
            candidates.append(u)
    for u in candidates:
        if u and not _is_bad_avatar(u):
            return u
    return ""


def get_ts(item):
    return g(item, "taken_at_timestamp", "taken_at", "legacy.created_at", "createdAt", "creation_time", "create_time", "createTime", "created_at", "createTimeISO",
             "publishedTimeSeconds", "publishedTimeStamp", "publish_time", "publishedAt",
             "publishedTime", "publishDate", "uploadDate", "timestamp", "date",
             "published_time_text", "publishedTimeText")


def to_wib(ts):
    if ts in (None, ""):
        return ("", "")
    try:
        if isinstance(ts, (int, float)) or (isinstance(ts, str) and ts.isdigit()):
            v = int(ts)
            if v > 10_000_000_000:  # milidetik
                v //= 1000
            dt = datetime.fromtimestamp(v, WIB)
        else:
            s = str(ts).strip()
            s = re.sub(r"\.\d+(?=Z$|[+-]\d{2}:?\d{2}$|$)", "", s)  # buang milidetik: .000Z -> Z
            try:
                dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(WIB)
            except Exception:
                dt = None
                for fmt in ("%a %b %d %H:%M:%S %z %Y", "%Y-%m-%dT%H:%M:%S%z",
                            "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
                    try:
                        dt = datetime.strptime(s, fmt)
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                        dt = dt.astimezone(WIB)
                        break
                    except Exception:
                        continue
                if dt is None:
                    rel = parse_relative(ts)
                    if rel is None:
                        return (str(ts), "")
                    dt = rel
        return (dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S"))
    except Exception:
        return (str(ts), "")


_DATE_ANY_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_TIME_ANY_RE = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")

def row_epoch_wib(date_s, time_s=""):
    """Epoch WIB dari kolom date/time dalam bentuk APA PUN (termasuk ISO mentah).
    None hanya kalau benar-benar tidak ada informasi tanggal."""
    ds = str(date_s or "")
    m = _DATE_ANY_RE.search(ds)
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    hh = mm = ss = 0
    mt = _TIME_ANY_RE.search(str(time_s or "")) or _TIME_ANY_RE.search(ds[m.end():])
    if mt:
        hh, mm, ss = int(mt.group(1)), int(mt.group(2)), int(mt.group(3) or 0)
    try:
        return int(datetime(y, mo, d, hh, mm, ss, tzinfo=WIB).timestamp())
    except Exception:
        return None

def build_url(platform, item, user, pid):
    u = g(item, "url", "share_url", "video_url", "permalink", "link")
    if u:
        return u
    code = g(item, "code", "shortcode")
    if platform.startswith("TikTok"):
        return f"https://www.tiktok.com/@{user}/video/{pid}" if pid else ""
    if platform.startswith("YouTube"):
        return f"https://www.youtube.com/watch?v={pid}" if pid else ""
    if platform.startswith("Threads"):
        return f"https://www.threads.net/@{user}/post/{code}" if code else ""
    if platform.startswith("Instagram"):
        return f"https://www.instagram.com/p/{code}" if code else ""
    return ""


_DRIVE = {"dir": None}

def ensure_drive():
    """Local version: output folder on the laptop (replaces Google Drive)."""
    if _DRIVE["dir"]:
        return _DRIVE["dir"]
    import os
    d = OUTPUT_DIR
    os.makedirs(d, exist_ok=True)
    _DRIVE["dir"] = d
    return d

CHECKPOINT_FILE = "checkpoint_partial.csv"
_CKPT = {"last": 0, "path": None}

def _checkpoint_path():
    try:
        if _DRIVE["dir"]:
            import os
            return os.path.join(_DRIVE["dir"], CHECKPOINT_FILE)
    except Exception:
        pass
    return CHECKPOINT_FILE

def save_checkpoint(rows, force=False):
    """Autosave berkala anti-disconnect: tulis progres ke Drive/lokal tiap ~50 baris."""
    n = len(rows)
    if not force and n - _CKPT["last"] < 50:
        return
    _CKPT["last"] = n
    try:
        import pandas as _pd
        p = _checkpoint_path()
        _pd.DataFrame(rows).to_csv(p, index=False)
        _CKPT["path"] = p
    except Exception:
        pass

SEEN_FILE = "seen_ids.json"

def _seen_path():
    try:
        if _DRIVE["dir"]:
            import os
            return os.path.join(_DRIVE["dir"], SEEN_FILE)
    except Exception:
        pass
    return SEEN_FILE

def load_seen_ids():
    try:
        with open(_seen_path()) as f:
            return set(_json.load(f))
    except Exception:
        return set()

def save_seen_ids(s):
    try:
        with open(_seen_path(), "w") as f:
            _json.dump(sorted(s), f)
    except Exception:
        pass

def check_credits():
    try:
        r = requests.get(f"{BASE}/v1/account/credit-balance", headers=HEADERS, timeout=20)
        if r.status_code == 200:
            return r.json()
        return {"status": r.status_code, "body": r.text[:120]}
    except Exception as e:
        return {"error": str(e)}


def _has_rt(rows):
    return any((r.get("type") == "retweet") or str(r.get("text", "")).startswith("RT @") for r in rows)

# Kata fungsi umum bahasa Indonesia (NETRAL topik) — untuk deteksi bahasa, bukan topik.
_ID_STOP = {"yang","dan","di","ini","itu","tidak","ada","dengan","untuk","dari","ke","pada","saya",
            "kami","kita","akan","sudah","belum","bisa","gak","nggak","aja","banget","kalau","karena",
            "atau","juga","mau","udah","lagi","sama","kok","sih","dong","deh","nya","pak","bu","dalam",
            "melalui","serta","agar","oleh","para","hingga","adalah","menjadi","tersebut","bagi","akan",
            "kepada","dalam","setelah","sebagai","namun","tetapi","bahwa","seperti","masih","sangat",
            "lebih","harus","dapat","telah","sudah","ketika","antara","hanya","sehingga","yaitu"}
_LATIN_RE = re.compile(r"[a-zA-Z]")

def _looks_indonesian(text):
    """Deteksi bahasa Indonesia berbasis kata fungsi netral + porsi aksara non-Latin.
    Ambang adaptif: teks pendek butuh 1 penanda, teks panjang butuh 2 (kurangi false-drop)."""
    t = str(text or "")
    if not t.strip():
        return True
    letters = _LATIN_RE.findall(t)
    non_ascii = [c for c in t if ord(c) > 0x2FF and not c.isspace()]
    if len(non_ascii) > max(6, 0.25 * max(1, len(t))):
        return False
    if len(letters) < 8:
        return True
    toks = re.findall(r"[a-z]+", t.lower())
    hits = len(set(toks) & _ID_STOP)
    need = 1 if len(toks) <= 12 else 2
    return hits >= need

LANG_FILTER = {"mode": ""}

def _lang_ok(text):
    m = LANG_FILTER.get("mode", "")
    if not m:
        return True
    if m == "id":
        return _looks_indonesian(text)
    if m == "en":
        t = str(text or "")
        non_ascii = [c for c in t if ord(c) > 0x2FF and not c.isspace()]
        if len(non_ascii) > max(6, 0.25 * max(1, len(t))):
            return False
        toks = set(re.findall(r"[a-z]+", t.lower()))
        en_stop = {"the","and","is","are","of","to","in","that","for","with","on","this","was","have"}
        return len(toks & en_stop) >= 2 or len(toks & _ID_STOP) < 2
    return True

# Mekanisme kata pengecualian (negative keywords) — GENERIK, tidak terikat topik apa pun.
# Default KOSONG: diisi manual oleh pengguna sesuai kampanye/klien masing-masing.
# (Contoh set siap-pakai untuk beberapa domain tersedia di file terpisah bila diperlukan.)
DEFAULT_NEG = []

def _text_has_negative(text, neg_terms):
    """True kalau teks memuat SALAH SATU kata pengecualian (substring, case-insensitive).
    Dipakai untuk MEMBUANG baris noise."""
    if not neg_terms:
        return False
    t = str(text or "").lower()
    for w in neg_terms:
        w = str(w).strip().lower()
        if w and w in t:
            return True
    return False

NEG_KW = {"terms": []}

def _text_matches_terms(text, terms):
    """True kalau teks memuat SALAH SATU frasa kata-kunci sebagai kata utuh (bukan substring).
    'balai wilayah sungai' cocok hanya bila ketiga kata itu muncul berurutan sebagai kata."""
    if not terms:
        return True
    t = " " + re.sub(r"[^\w\s]", " ", str(text or "").lower()) + " "
    t = re.sub(r"\s+", " ", t)
    for phrase in terms:
        ph = re.sub(r"[^\w\s]", " ", str(phrase).lower()).strip()
        ph = re.sub(r"\s+", " ", ph)
        if not ph:
            continue
        if f" {ph} " in t:
            return True
    return False

def _rt_original(text):
    m = re.match(r"RT @([A-Za-z0-9_]+):", str(text or ""))
    return m.group(1).lower() if m else None

def _cap_retweets(rows, log):
    """Batasi jumlah retweet: per-tweet-asli dan/atau total. Non-RT tidak tersentuh."""
    per = TW_INC_RT.get("cap_per_orig", 0) or 0
    tot = TW_INC_RT.get("cap_total", 0) or 0
    if per <= 0 and tot <= 0:
        return rows
    out, per_count, total_rt, dropped = [], {}, 0, 0
    for r in rows:
        is_rt = r.get("type") == "retweet" or str(r.get("text", "")).startswith("RT @")
        if not is_rt:
            out.append(r); continue
        if tot and total_rt >= tot:
            dropped += 1; continue
        if per:
            key = _rt_original(r.get("text")) or "?"
            if per_count.get(key, 0) >= per:
                dropped += 1; continue
            per_count[key] = per_count.get(key, 0) + 1
        total_rt += 1; out.append(r)
    if dropped:
        lim = []
        if per: lim.append(f"maks {per}/tweet-asli")
        if tot: lim.append(f"maks {tot} total")
        log(f"  Twitter: {dropped} retweet(s) dropped by limits ({', '.join(lim)}); {total_rt} RT kept")
    return out

def tw_search_with_retweets(tterm, max_items, want_raw, log):
    """Pastikan retweet ikut tertangkap:
    1) coba include:nativeretweets;
    2) kalau hasilnya tanpa satu pun RT (operator diabaikan provider), tarik stream KHUSUS retweet
       via filter:nativeretweets lalu gabungkan (dedup per ID)."""
    rows = scrape_twitter_getx(tterm + " include:nativeretweets", max_items, want_raw, log)
    if rows and _has_rt(rows):
        return _cap_retweets(rows, log)  # operator dihormati -> selesai, 1x biaya
    if not rows:
        log("  Twitter: 0 results with include: \u2014 retrying without the operator...")
        rows = scrape_twitter_getx(tterm, max_items, want_raw, log)
    if _has_rt(rows):
        return _cap_retweets(rows, log)
    log("  Twitter: no retweets in results \u2014 fetching retweets via filter:nativeretweets...")
    rt_rows = scrape_twitter_getx(tterm + " filter:nativeretweets", max_items, want_raw, log)
    seen_ids = {str(r.get("id")) for r in rows}
    extra = [r for r in rt_rows if str(r.get("id")) not in seen_ids]
    for r in extra:
        if r.get("type") not in ("retweet",):
            r["type"] = "retweet"
    rows = rows + extra
    rows = _cap_retweets(rows, log)
    extra = [r for r in rows if r.get("type") == "retweet"]
    if extra:
        log(f"  Twitter: {len(extra)} retweet(s) after limits")
    else:
        log("  Twitter: the provider returned no retweets for either operator \u2014 likely an API limitation.")
    return rows + extra


def scrape(platform, term, max_items, since_d, until_d, want_raw, log, s_ep=None, u_ep=None):
    if platform.startswith("Facebook"):
        return scrape_fb_via_google(term, max_items, want_raw, log, s_ep=s_ep, u_ep=u_ep)
    if platform == IG_KW_PLATFORM:
        return scrape_ig_via_google(term, max_items, want_raw, log, s_ep=s_ep, u_ep=u_ep)
    if platform.startswith("Twitter"):
        tterm = term
        if TW_LANG["code"] and "lang:" not in tterm:
            tterm += f" lang:{TW_LANG['code']}"
        if since_d: tterm += f" since:{since_d}"
        if until_d: tterm += f" until:{until_d}"
        if TW_INC_RT["on"] and "nativeretweets" not in tterm:
            return tw_search_with_retweets(tterm, max_items, want_raw, log)
        return scrape_twitter_getx(tterm, max_items, want_raw, log)
    cfg = PLATFORMS[platform]
    base_params = {cfg["param"]: term}
    if cfg.get("dates"):
        if since_d: base_params["start_date"] = since_d
        if until_d: base_params["end_date"] = until_d

    rows, seen, cursor, page, retries = [], set(), None, 0, 0
    while len(rows) < max_items and page < 150:
        page_dated, page_inrange = 0, 0
        params = dict(base_params)
        if cursor:
            ckey = "continuationToken" if "youtube" in cfg["path"] else "cursor"
            params[ckey] = cursor
        try:
            r = requests.get(f"{BASE}{cfg['path']}", headers=HEADERS, params=params, timeout=45)
        except Exception as e:
            log(f"  connection error: {e}"); break
        if r.status_code != 200:
            body = r.text[:160]
            if r.status_code == 404 or "not_found" in body or "No posts" in body:
                log("  done (no more results)"); break
            if ("429" in body or r.status_code in (429, 500)) and retries < 5:
                retries += 1
                wait = 5 * retries
                log(f"  rate-limited, waiting {wait}s then retrying (retry {retries}/5)..."); time.sleep(wait); continue
            log(f"  Error {r.status_code}: {body}"); break
        data = r.json()
        items = find_items(data)
        if page == 0:
            log(f"  (response: keys={list(data.keys()) if isinstance(data, dict) else 'list'}, items found={len(items)})")
        if not items:
            break
        new = 0
        for it_raw in items:
            it = unwrap(it_raw)
            pid = get_id(it)
            dedup = pid or get_text(it)[:60]
            if dedup in seen:
                continue
            seen.add(dedup)
            user = get_user(it)
            tgl, jam = to_wib(get_ts(it))
            if (s_ep or u_ep) and not cfg.get("dates"):
                _pe = row_epoch_wib(tgl, jam)
                if _pe is not None:
                    page_dated += 1
                    if (not s_ep or _pe >= s_ep) and (not u_ep or _pe < u_ep):
                        page_inrange += 1
            row = {
                "platform": platform.split(" ")[0],
                "type": media_subtype(platform, it, g(it, "url", "share_url", "permalink") or ""),
                "date": tgl,
                "time": jam,
                "username": user, "user_id": get_user_id(it), "profile_image": get_profile_image(it),
                "text": get_text(it).replace("\n", " "),
                "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite"]),
                "comments": num_at(it, COMMENT_PATHS, ["comment", "reply"]),
                "shares": num_at(it, SHARE_PATHS, ["share", "repost", "reshare"]),
                "views": num_at(it, VIEW_PATHS, ["view", "play"]),
                "followers": get_followers(it),
                "id": pid,
                "url": build_url(platform, it, user, pid),
                "query": term,
            }
            if want_raw:
                row["raw_json"] = _json.dumps(it_raw, ensure_ascii=False)[:2000]
            rows.append(row)
            new += 1
            if len(rows) >= max_items:
                break
        page += 1
        if s_ep and not cfg.get("dates") and page_dated >= 5 and page_inrange == 0:
            log(f"  {platform}: a full page is outside the time range -> stopping pagination (saves credits), {len(rows)} rows")
            break
        cursor = None
        if isinstance(data, dict):
            cursor = (data.get("cursor") or data.get("next_cursor") or data.get("max_id")
                      or data.get("nextCursor") or data.get("paging_token")
                      or data.get("continuationToken") or data.get("continuation_token")
                      or data.get("next_page_token") or data.get("end_cursor"))
        log(f"  ...{len(rows)} items")
        if not cursor or new == 0:
            break
        time.sleep(0.2)
    return rows


TW_INC_RT = {"on": True, "cap_per_orig": 0, "cap_total": 0}
STRICT_KW = {"on": False, "terms": []}
TW_LANG = {"code": ""}
LINK_MODE = "BY LINK (paste URLs)"
PROFILE_MODE = "BY PROFILE (paste profile URLs)"
MIXED_MODE = "MIXED (keyword + link + profil sekaligus)"

_POST_URL_PAT = re.compile(r"(/status/|/watch\?|youtu\.be/|/shorts/|/video/|/reel/|/reels/|/p/|/post/|/posts/|/photo|/share/p/|/share/v/|/share/r/)", re.I)
def looks_like_post_url(u):
    """True = link ke postingan; False = link profil/akun."""
    return bool(_POST_URL_PAT.search(u or ""))

DETAIL_EP = {
    "TikTok":    "/v2/tiktok/video",
    "Facebook":  "/v1/facebook/post",
    "Threads":   "/v1/threads/post",
    "Instagram": "/v1/instagram/post",
    "YouTube":   "/v1/youtube/video",
    "Twitter":   "/v1/twitter/tweet",
}

def detect_platform(url):
    u = url.lower()
    if "tiktok.com" in u: return "TikTok"
    if "youtube.com" in u or "youtu.be" in u: return "YouTube"
    if "threads.com" in u or "threads.net" in u: return "Threads"
    if "instagram.com" in u: return "Instagram"
    if "facebook.com" in u or "fb.watch" in u: return "Facebook"
    if "x.com/" in u or "twitter.com" in u: return "Twitter"
    return None

def yt_id(url):
    for pat in [r"[?&]v=([\w-]+)", r"youtu\.be/([\w-]+)", r"/shorts/([\w-]+)"]:
        m = re.search(pat, url)
        if m:
            return m.group(1)
    return ""

DESCEND_KEYS = ["data", "xdt_shortcode_media", "aweme_detail", "result", "tweet", "post", "video_info"]

def pick_post(data):
    cur = data
    for _ in range(6):
        cur = unwrap(cur)
        if not isinstance(cur, dict):
            break
        moved = False
        for k in DESCEND_KEYS:
            if isinstance(cur.get(k), dict):
                cur = cur[k]; moved = True; break
        if not moved:
            break
    return cur

def is_fb_post_url(u):
    u = u.lower()
    if "facebook.com" not in u:
        return False
    return any(k in u for k in ["/posts/", "/reel/", "/videos/", "/video/", "/watch", "story.php", "permalink", "/share/"])

def google_search(query, page, log):
    for attempt in range(4):
        try:
            r = requests.get(BASE + "/v1/google/search", headers=HEADERS,
                             params={"query": query, "page": page}, timeout=40)
        except Exception as e:
            log(f"  google error: {e}"); return None
        if r.status_code == 200:
            return r.json()
        body = r.text[:90]
        if ("429" in body or r.status_code in (429, 500)) and attempt < 3:
            wait = 5 * (attempt + 1)
            log(f"  google rate-limited, waiting {wait}s (retry {attempt+1}/3)..."); time.sleep(wait); continue
        log(f"  google HTTP {r.status_code}: {body}"); return None
    return None

_NATIVE_KW = {}   # platform -> {"path":..., "param":..., "qkey":...} atau False (tak ada)

def _native_keyword_search(platform, ep_candidates, term, max_items, want_raw, log, s_ep=None, u_ep=None):
    """Coba endpoint keyword/search NATIVE provider (bukan Google). Jauh lebih lengkap & baru.
    ep_candidates: list (path, nama_param_query). Return rows atau None kalau tak ada yang jalan."""
    key = platform
    if key in _NATIVE_KW and _NATIVE_KW[key] is False:
        return None
    cands = ([( _NATIVE_KW[key]["path"], _NATIVE_KW[key]["param"] )]
             if _NATIVE_KW.get(key) else ep_candidates)
    if key not in _NATIVE_KW:
        log(f"  \U0001F50D {platform}: probing the provider's native keyword endpoint...")
    for path, qp in cands:
        try:
            r = requests.get(BASE + path, headers=HEADERS, params={qp: term, "count": max_items}, timeout=40)
        except BudgetExceeded:
            raise
        except Exception:
            continue
        if r.status_code != 200:
            if key not in _NATIVE_KW:
                log(f"     {path} -> HTTP {r.status_code}")
            continue
        items = find_items(r.json())
        if not items:
            if key not in _NATIVE_KW:
                log(f"     {path} -> 200 but 0 items")
            continue
        _NATIVE_KW[key] = {"path": path, "param": qp}
        log(f"  \u2705 {platform}: using native keyword endpoint '{path}' ({len(items)} items on page 1).")
        rows = []
        pshort = platform.split(" ")[0]
        for it_raw in items[:max_items]:
            it = unwrap(it_raw)
            tgl, jam = to_wib(get_ts(it))
            u = str(g(it, "url", "share_url", "permalink", "webVideoUrl", "link") or "")
            rows.append({
                "platform": pshort, "type": media_subtype(pshort, it, u),
                "date": tgl, "time": jam, "username": get_user(it),
                "user_id": get_user_id(it), "profile_image": get_profile_image(it),
                "_profile": prof_ident(pshort, it, get_user(it)),
                "text": get_text(it).replace("\n", " "),
                "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite", "reaction"]),
                "comments": num_at(it, COMMENT_PATHS, ["comment", "reply"]),
                "shares": num_at(it, SHARE_PATHS, ["share", "repost", "reshare"]),
                "views": num_at(it, VIEW_PATHS, ["view", "play"]),
                "followers": get_followers(it),
                "id": get_id(it), "url": u, "query": term,
            })
            if want_raw:
                rows[-1]["raw_json"] = _json.dumps(it_raw, ensure_ascii=False)[:2000]
        return rows
    if key not in _NATIVE_KW:
        _NATIVE_KW[key] = False
        log(f"  \u2139\uFE0F {platform}: no native keyword endpoint \u2014 using Google (limited index).")
    return None

def scrape_fb_via_google(term, max_items, want_raw, log, s_ep=None, u_ep=None):
    _native = _native_keyword_search("Facebook", [
        ("/v1/facebook/search/posts", "query"), ("/v1/facebook/search", "query"),
        ("/v1/facebook/posts/search", "keyword"), ("/v2/facebook/search/posts", "q")],
        term, max_items, want_raw, log, s_ep, u_ep)
    if _native is not None:
        return _native
    queries = [f"{term} site:facebook.com"]
    if COVERAGE["on"]:
        queries += [f'"{term}" site:facebook.com',
                    f"{term} site:facebook.com/reel",
                    f"{term} site:facebook.com/watch"]
    fb_urls = _google_collect_urls(queries, is_fb_post_url, max_items,
                                   10 if COVERAGE["on"] else 8, log, "FB")

    def _fb_one(u):
        try:
            r = requests.get(BASE + "/v1/facebook/post", headers=HEADERS, params={"url": u}, timeout=45)
        except Exception as e:
            log(f"  fb error: {e}"); return None
        if r.status_code != 200:
            log(f"  fb HTTP {r.status_code}"); return None
        d = r.json()
        it = pick_post(d)
        tgl, jam = to_wib(get_ts(it))
        user = get_user(it); pid = get_id(it)
        row = {
            "platform": "Facebook", "type": media_subtype("Facebook", it, u),
            "date": tgl, "time": jam, "username": user, "user_id": get_user_id(it), "profile_image": get_profile_image(it),
            "_profile": prof_ident("Facebook", it, user),
            "text": get_text(it).replace("\n", " "),
            "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite", "reaction"]),
            "comments": num_at(it, COMMENT_PATHS, ["comment", "reply"]),
            "shares": num_at(it, SHARE_PATHS, ["share", "repost", "reshare"]),
            "views": num_at(it, VIEW_PATHS, ["view", "play"]),
            "followers": get_followers(it),
            "id": pid, "url": u, "query": term,
        }
        if want_raw:
            row["raw_json"] = _json.dumps(d, ensure_ascii=False)[:2000]
        return row
    pool = fb_urls[:max_items]
    if not (s_ep or u_ep):
        rows = [r for r in parallel_map(_fb_one, pool, workers=8) if r]
        log(f"  Facebook: {len(rows)} posts fetched")
        return rows
    rows, got_dated, total_inr = [], 0, 0
    for k in range(0, len(pool), 10):
        chunk = [r for r in parallel_map(_fb_one, pool[k:k+10], workers=8) if r]
        chunk_inr = 0
        for row in chunk:
            _pe = row_epoch_wib(row.get("date"), row.get("time"))
            if _pe is None:
                rows.append(row); continue
            got_dated += 1
            if (s_ep and _pe < s_ep) or (u_ep and _pe >= u_ep):
                continue
            chunk_inr += 1; total_inr += 1; rows.append(row)
        if got_dated >= 10 and chunk_inr == 0 and total_inr == 0:
            log("  Facebook: first 10+ links are all outside the time range -> stopping (saves credits)")
            break
    log(f"  Facebook: {len(rows)} posts in range")
    return rows

GETX_BASE = _b64.b64decode("aHR0cHM6Ly9hcGkuZ2V0eGFwaS5jb20=").decode()

def _tweet_id(url):
    m = re.search(r"status/(\d+)", url)
    return m.group(1) if m else ""

def _tw_rt_obj(it):
    """Ambil objek tweet asli dari sebuah retweet (berbagai bentuk nesting API)."""
    rt = g(it, "retweeted_status", "retweeted_tweet",
           "legacy.retweeted_status_result.result", "legacy.retweeted_status_result",
           "retweeted_status_result.result", "retweeted_status_result")
    if isinstance(rt, dict):
        rt = unwrap(rt)
        if isinstance(rt, dict) and isinstance(rt.get("tweet"), dict):
            rt = rt["tweet"]
        return rt if isinstance(rt, dict) else None
    return None

def tw_full_text(it, txt):
    """Retweet sering terpotong 'RT @x: ...' — susun ulang teks lengkap dari tweet asli."""
    rt = _tw_rt_obj(it)
    if rt and (not txt or txt.rstrip().endswith(("\u2026", "..."))):
        ou = get_user(rt)
        ot = get_text(rt).replace("\n", " ")
        if ot:
            txt = f"RT @{ou}: {ot}" if ou else ot
    return txt, rt

def tw_metric(it, rt, paths, kws):
    v = num_at(it, paths, kws)
    if rt is not None and v in (None, "", 0):
        v2 = num_at(rt, paths, kws)
        if v2 not in (None, ""):
            return v2
    return v

def _fetch_tweet_getx(url, want_raw, log):
    tid = _tweet_id(url)
    if not tid:
        log(f"  {url[:48]} -> tweet id not found"); return None
    try:
        r = requests.get(GETX_BASE + "/twitter/tweet/detail", params={"id": tid},
                         headers={"Authorization": f"Bearer {TWITTER_API_KEY.strip()}"}, timeout=45)
    except Exception as e:
        log(f"  Twitter error: {e}"); return None
    if r.status_code != 200:
        log(f"  Twitter HTTP {r.status_code}: {r.text[:90]} -> retrying once...")
        time.sleep(2)
        try:
            r = requests.get(GETX_BASE + "/twitter/tweet/detail", params={"id": tid},
                             headers={"Authorization": f"Bearer {TWITTER_API_KEY.strip()}"}, timeout=45)
        except Exception as e:
            log(f"  Twitter error: {e}"); return None
        if r.status_code != 200:
            log(f"  Twitter HTTP {r.status_code} (final): {r.text[:90]}")
            log("  Note: some retweet IDs cannot be looked up directly; try pasting the ORIGINAL tweet URL.")
            return None
    data = r.json()
    it = pick_post(data)
    if not get_id(it):
        items = [unwrap(x) for x in find_items(data)]
        match = [x for x in items if get_id(x) == tid]
        if match:
            it = match[0]
        elif items:
            it = items[0]
    tgl, jam = to_wib(get_ts(it))
    user = get_user(it); pid = get_id(it) or tid
    if not user or " " in user:
        user = _handle_from_url(url) or user   # x.com/<handle>/status/... pasti berisi handle
    _txt = get_text(it).replace("\n", " ")
    _txt, _rt = tw_full_text(it, _txt)
    row = {
        "platform": "Twitter", "type": tweet_subtype(it, _txt),
        "date": tgl, "time": jam, "username": user, "user_id": get_user_id(it), "profile_image": get_profile_image(it),
        "text": _txt,
        "likes": tw_metric(it, _rt, LIKE_PATHS, ["like", "favorite"]),
        "comments": tw_metric(it, _rt, COMMENT_PATHS, ["comment", "reply"]),
        "shares": tw_metric(it, _rt, SHARE_PATHS, ["share", "repost", "reshare", "retweet"]),
        "views": tw_metric(it, _rt, VIEW_PATHS, ["view", "play"]),
        "followers": get_followers(it),
        "id": pid, "url": url, "query": "(by link)",
    }
    if _rt and not row["username"]:
        row["username"] = get_user(_rt); row["user_id"] = row["user_id"] or get_user_id(_rt)
        row["profile_image"] = row["profile_image"] or get_profile_image(_rt)
    if want_raw:
        row["raw_json"] = _json.dumps(data, ensure_ascii=False)[:2000]
    return row

COVERAGE = {"on": False}

def scrape_twitter_getx(term, max_items, want_raw, log):
    if not TWITTER_API_KEY.strip():
        log("  Twitter keyword search needs a Twitter API key (fill it in the form)."); return []
    products = ["Latest", "Top"] if COVERAGE["on"] else ["Latest"]
    out, seen_ids = [], set()
    for prod in products:
        for r0 in _tw_search_one(term, prod, max_items, want_raw, log):
            k = r0.get("id") or r0.get("text", "")[:60]
            if k in seen_ids:
                continue
            seen_ids.add(k)
            out.append(r0)
            if len(out) >= max_items and prod == products[-1]:
                break
    return out

def _tw_search_one(term, product, max_items, want_raw, log):
    rows, cursor, page = [], None, 0
    while len(rows) < max_items and page < 80:
        params = {"q": term, "product": product}
        if cursor:
            params["cursor"] = cursor
        try:
            r = requests.get(GETX_BASE + "/twitter/tweet/advanced_search", params=params,
                             headers={"Authorization": f"Bearer {TWITTER_API_KEY.strip()}"}, timeout=45)
        except Exception as e:
            log(f"  Twitter error: {e}"); break
        if r.status_code != 200:
            log(f"  Twitter HTTP {r.status_code}: {r.text[:90]}"); break
        data = r.json()
        items = find_items(data)
        if page == 0:
            log(f"  (response keys={list(data.keys()) if isinstance(data, dict) else 'list'}, items found={len(items)})")
        if not items:
            break
        for it_raw in items:
            it = unwrap(it_raw)
            _txt = get_text(it).replace("\n", " ")
            _txt, _rt = tw_full_text(it, _txt)
            row = {
                "platform": "Twitter", "type": tweet_subtype(it, _txt),
                "date": to_wib(get_ts(it))[0], "time": to_wib(get_ts(it))[1],
                "username": get_user(it), "user_id": get_user_id(it), "profile_image": get_profile_image(it),
                "text": _txt,
                "likes": tw_metric(it, _rt, LIKE_PATHS, ["like", "favorite"]),
                "comments": tw_metric(it, _rt, COMMENT_PATHS, ["comment", "reply"]),
                "shares": tw_metric(it, _rt, SHARE_PATHS, ["share", "repost", "reshare", "retweet"]),
                "views": tw_metric(it, _rt, VIEW_PATHS, ["view", "play"]),
                "followers": get_followers(it),
                "id": get_id(it), "url": g(it, "url") or "", "query": term,
            }
            if want_raw:
                row["raw_json"] = _json.dumps(it_raw, ensure_ascii=False)[:2000]
            rows.append(row)
            if len(rows) >= max_items:
                break
        page += 1
        cursor = None
        if isinstance(data, dict):
            cursor = data.get("next_cursor") or data.get("cursor") or data.get("nextCursor")
        log(f"  ...{len(rows)} items")
        if not cursor:
            break
        time.sleep(0.3)
    return rows

def fetch_tw_replies(post_id, want_raw, log, cap=50):
    """Ambil reply sebuah tweet via advanced_search conversation_id: (getxapi)."""
    if not TWITTER_API_KEY.strip() or not post_id:
        return []
    n = cap if cap and cap > 0 else 200
    rows = _tw_search_one(f"conversation_id:{post_id}", "Latest", n, want_raw, log)
    out = []
    for r in rows:
        if str(r.get("id")) == str(post_id):   # buang tweet induknya sendiri
            continue
        r["type"] = "comment"
        r["parent_id"] = str(post_id)
        r["query"] = "(reply)"
        out.append(r)
    return out

PROFILE_POSTS_EP = {
    "Instagram": ("/v2/instagram/user/posts",   "handle"),
    "TikTok":    ("/v3/tiktok/profile/videos",  "handle"),
    "YouTube":   ("/v1/youtube/channel-videos", "handle"),
    "Facebook":  ("/v1/facebook/profile/posts", "url"),
    "Threads":   ("/v1/threads/user/posts",     "handle"),
}
def _recover_pagination(full_url, base_params, cursor_val, seen_ids, page, hint_skip=None):
    """Halaman berulang: jajaki nama parameter pagination alternatif (maks 8 request).
    Return (mode, name) bila ada varian yang menghasilkan item BARU; None bila semua gagal."""
    tries = []
    if cursor_val is not None:
        for name in ("cursor", "max_cursor", "next_cursor", "next_max_id",
                     "pagination_token", "page_token", "after", "end_cursor"):
            if name != hint_skip:
                tries.append(("cursor", name, {name: cursor_val}))
    tries.append(("page", "page", {"page": page + 1}))
    tries.append(("offset", "offset", {"offset": max(1, len(seen_ids))}))
    for mode, name, extra in tries[:8]:
        pp = dict(base_params); pp.update(extra)
        try:
            r = requests.get(full_url, headers=HEADERS, params=pp, timeout=45)
            if r.status_code != 200:
                continue
            batch = find_items(r.json())
        except BudgetExceeded:
            raise
        except Exception:
            continue
        for b in batch or []:
            bu = unwrap(b)
            bid = get_id(bu) or get_text(bu)[:80]
            if bid and bid not in seen_ids:
                return (mode, name)
    return None

CURSOR_PARAM = {"TikTok": "max_cursor", "Instagram": "cursor", "YouTube": "continuationToken",
                "Facebook": "cursor", "Threads": "cursor"}

def profile_handle(url, plat):
    u = url.strip().rstrip("/")
    if plat == "YouTube":
        m = re.search(r"/channel/([A-Za-z0-9_-]+)", u)
        if m: return m.group(1)
        m = re.search(r"@([A-Za-z0-9._-]+)", u)
        if m: return m.group(1)
        m = re.search(r"youtube\.com/(?:c/|user/)?([A-Za-z0-9._-]+)", u)
        return m.group(1) if m else ""
    m = re.search(r"@([A-Za-z0-9._-]+)", u)
    if m: return m.group(1)
    m = re.search(r"(?:instagram\.com|threads\.(?:net|com)|tiktok\.com|x\.com|twitter\.com|facebook\.com)/([A-Za-z0-9._-]+)", u)
    if m:
        h = m.group(1)
        if h.lower() not in ("channel", "c", "user", "p", "reel", "reels", "share", "watch"):
            return h
    return ""

def _find_cursor(data):
    if not isinstance(data, dict):
        return None
    for k in ("max_cursor", "next_max_id", "cursor", "next_cursor", "nextCursor",
              "continuationToken", "end_cursor", "paging_token"):
        v = data.get(k)
        if v and str(v) not in ("0", "0.0", "false", "None", ""):
            return (k, v)
    return None

def scrape_profile(url, max_items, want_raw, log, since_d="", until_d="", s_ep=None, u_ep=None):
    plat = detect_platform(url)
    if plat is None:
        log(f"  {url[:48]} -> unknown platform"); return []
    if plat == "Twitter":
        h = profile_handle(url, plat)
        if not h:
            log("  Twitter handle not found"); return []
        q = f"from:{h}"
        if since_d: q += f" since:{since_d}"
        if until_d: q += f" until:{until_d}"
        if TW_INC_RT["on"]:
            return tw_search_with_retweets(q, max_items, want_raw, log)
        return scrape_twitter_getx(q, max_items, want_raw, log)
    ep = PROFILE_POSTS_EP.get(plat)
    if not ep:
        log(f"  {plat} profile-posts not supported"); return []
    path, pkey = ep
    ident = url.strip() if pkey == "url" else profile_handle(url, plat)
    if not ident:
        log(f"  {plat}: handle not found in URL"); return []
    params = {pkey: ident}
    if plat == "YouTube" and ident.startswith("UC") and len(ident) > 20:
        params = {"channelId": ident}
    rows, page, cursor_kv = [], 0, None
    seen_pids = set()
    pag_mode, pag_name, tried_recovery = "cursor", None, False
    oldest_ep = None            # epoch post tertua yang sudah didapat (untuk date-window paging)
    date_window_active = False
    window_until = u_ep
    while len(rows) < max_items and page < 80:
        new_this_page = 0
        page_dated, page_older = 0, 0
        pp = dict(params)
        if pag_mode == "page":
            pp[pag_name] = page + 1
        elif pag_mode == "offset":
            pp[pag_name] = len(seen_pids)
        elif pag_mode == "datewindow":
            from datetime import datetime as _dtm
            _ds = _dtm.fromtimestamp(window_until, WIB).strftime("%Y-%m-%d")
            for _pn in ("until", "max_date", "end_date", "before", "until_date", "to"):
                pp[_pn] = _ds     # kirim beberapa nama umum; provider pakai yang dikenalinya
        elif cursor_kv:
            pp[pag_name or CURSOR_PARAM.get(plat, "cursor")] = cursor_kv[1]
        try:
            r = requests.get(BASE + path, headers=HEADERS, params=pp, timeout=45)
        except Exception as e:
            log(f"  {plat} error: {e}"); break
        if r.status_code != 200:
            log(f"  {plat} HTTP {r.status_code}: {r.text[:90]}"); break
        data = r.json()
        items = find_items(data)
        if page == 0:
            log(f"  {plat}: keys={list(data.keys()) if isinstance(data, dict) else 'list'}, items={len(items)}")
        if not items:
            break
        for it_raw in items:
            it = unwrap(it_raw)
            user = get_user(it); pid = get_id(it)
            pkey = pid or get_text(it)[:80]
            if pkey in seen_pids:      # API mengulang item yang sama -> lewati
                continue
            seen_pids.add(pkey); new_this_page += 1
            tgl, jam = to_wib(get_ts(it))
            if (s_ep or u_ep) and tgl:
                _ep = row_epoch_wib(tgl, jam)
                if _ep is not None:
                    page_dated += 1
                    if u_ep and _ep >= u_ep:
                        continue                      # lebih baru dari rentang -> lewati
                    if s_ep and _ep < s_ep:
                        page_older += 1; continue     # lebih tua dari rentang -> lewati
            row = {
                "platform": plat, "type": media_subtype(plat, it, g(it, "url", "share_url", "permalink") or ""),
                "date": tgl, "time": jam,
                "username": user, "user_id": get_user_id(it), "profile_image": get_profile_image(it),
                "text": get_text(it).replace("\n", " "),
                "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite", "reaction"]),
                "comments": num_at(it, COMMENT_PATHS, ["comment", "reply"]),
                "shares": num_at(it, SHARE_PATHS, ["share", "repost", "reshare", "retweet"]),
                "views": num_at(it, VIEW_PATHS, ["view", "play"]),
                "followers": get_followers(it),
                "id": pid, "url": g(it, "url", "share_url", "permalink") or "", "query": f"profile:{ident}",
                "_profile": prof_ident(plat, it, user),
            }
            if want_raw:
                row["raw_json"] = _json.dumps(it_raw, ensure_ascii=False)[:2000]
            rows.append(row)
            _re2 = row_epoch_wib(tgl, jam)
            if _re2 is not None and (oldest_ep is None or _re2 < oldest_ep):
                oldest_ep = _re2
            if len(rows) >= max_items:
                break
        page += 1
        if new_this_page == 0:
            # Timeframe diset -> paling andal: paging berbasis tanggal (mundurkan 'until')
            if (s_ep or u_ep) and oldest_ep and (not s_ep or oldest_ep > s_ep) and pag_mode != "datewindow_exhausted":
                nu = oldest_ep - 1
                if window_until and nu >= window_until:
                    nu = window_until - 3600
                if not s_ep or nu >= s_ep:
                    window_until = nu
                    pag_mode, pag_name, cursor_kv = "datewindow", None, None
                    from datetime import datetime as _dtm
                    log(f"  {plat}: cursor repeated -> continuing via DATE window up to {_dtm.fromtimestamp(window_until, WIB):%Y-%m-%d} (follows the time range)")
                    continue
            if not tried_recovery:
                tried_recovery = True
                cur0 = _find_cursor(data)
                rec = _recover_pagination(BASE + path, params, (cur0 or cursor_kv or (None, None))[1],
                                          seen_pids, page,
                                          hint_skip=pag_name or CURSOR_PARAM.get(plat, "cursor"))
                if rec:
                    pag_mode, pag_name = rec
                    if cur0:
                        cursor_kv = cur0
                    log(f"  {plat}: pagination recovered via parameter '{pag_name}' -> continuing")
                    continue
            if pag_mode == "datewindow":
                pag_mode = "datewindow_exhausted"   # jendela tanggal pun tak memberi post baru -> benar2 habis
            if s_ep or u_ep:
                log(f"  {plat}: stopped at {len(rows)} posts \u2014 everything up to the time-range boundary was fetched (or the account has no older posts in range).")
            else:
                log(f"  {plat}: provider repeated the page and all pagination methods failed -> stopped at {len(rows)} posts")
                log(f"    (tip: enable the time filter to continue via date paging; endpoint {path} ignores the cursor)")
            break
        if s_ep and page_dated > 0 and page_older == page_dated:
            log(f"  {plat}: whole page is older than the time range -> stopping pagination (saves credits), {len(rows)} posts"); break
        cur = _find_cursor(data)
        more = None
        if isinstance(data, dict):
            more = data.get("has_more", data.get("more_available", data.get("has_next_page")))
        if more is not None and not more:
            break
        if pag_mode in ("page", "offset"):
            time.sleep(0.2)
            continue
        if not cur:
            break
        if cursor_kv and str(cur[1]) == str(cursor_kv[1]):
            # cursor mengulang. Timeframe diset -> geser jendela tanggal (paling andal).
            if (s_ep or u_ep) and oldest_ep and (not s_ep or oldest_ep > s_ep) and pag_mode != "datewindow_exhausted":
                nu = oldest_ep - 1
                if window_until and nu >= window_until:
                    nu = window_until - 3600
                if not s_ep or nu >= s_ep:
                    window_until = nu
                    pag_mode, pag_name, cursor_kv = "datewindow", None, None
                    from datetime import datetime as _dtm
                    log(f"  {plat}: cursor repeated -> continuing via DATE window up to {_dtm.fromtimestamp(window_until, WIB):%Y-%m-%d} (follows the time range)")
                    time.sleep(0.2); continue
            if not tried_recovery:
                tried_recovery = True
                rec = _recover_pagination(BASE + path, params, cur[1], seen_pids, page,
                                          hint_skip=pag_name or CURSOR_PARAM.get(plat, "cursor"))
                if rec:
                    pag_mode, pag_name = rec
                    log(f"  {plat}: cursor unchanged but pagination recovered via '{pag_name}' -> continuing")
                    cursor_kv = cur
                    time.sleep(0.2)
                    continue
            if s_ep or u_ep:
                log(f"  {plat}: {len(rows)} posts \u2014 the whole time range is covered."); break
            log(f"  {plat}: cursor unchanged and other variants failed -> stopped at {len(rows)} posts"); break
        cursor_kv = cur
        time.sleep(0.2)
    log(f"  {plat}: {len(rows)} posts")
    return rows

COMMENTS_EP = {
    "TikTok":    "/v1/tiktok/video/comments",
    "YouTube":   "/v1/youtube/video/comments",
    "Instagram": "/v2/instagram/post/comments",
    "Facebook":  "/v1/facebook/post/comments",
}

def _threads_clean_url(u):
    return (u or "").split("?")[0].split("#")[0].rstrip("/")

def _threads_code(u):
    m = re.search(r"/post/([A-Za-z0-9_-]+)", u or "")
    return m.group(1) if m else ""

_THREADS_CMT = {"path": None, "kind": None, "checked": False}

def _threads_param(kind, post_url, post_id):
    cu = _threads_clean_url(post_url)
    if kind == "url":
        return {"url": cu}
    if kind == "code":
        c = _threads_code(cu)
        return {"id": c} if c else None
    if kind == "post_id":
        return {"id": str(post_id)} if post_id else None
    if kind == "pk":
        pk = str(post_id or "").split("_")[0]
        return {"id": pk} if pk else None
    return None

def _threads_cmt_ep(post_url, post_id=None):
    """Jajaki endpoint komentar Threads dengan semua identitas yang tersedia:
    URL bersih, shortcode (/post/xxx), ID post penuh, dan pk (angka sebelum '_')."""
    if _THREADS_CMT["checked"]:
        return _THREADS_CMT["path"]
    _THREADS_CMT["checked"] = True
    kinds = []
    for k in ("url", "code", "post_id", "pk"):
        p = _threads_param(k, post_url, post_id)
        if p and all(v for v in p.values()) and (k != "pk" or p["id"] != str(post_id)):
            kinds.append((k, p))
    paths = ("/v1/threads/post/replies", "/v1/threads/post/comments", "/v2/threads/post/comments")
    total = len(paths) * len(kinds)
    print(f"  \U0001F50D Threads: probing the comments endpoint (once per session, max {total} requests)...")
    for path in paths:
        for kind, params in kinds:
            try:
                r = requests.get(BASE + path, headers=HEADERS, params=params, timeout=30)
                n = len(find_items(r.json())) if r.status_code == 200 else 0
                pv = list(params.values())[0]
                print(f"     {path} [{kind}={str(pv)[:28]}] -> HTTP {r.status_code}" +
                      (f", {n} item" if r.status_code == 200 else ""))
                if r.status_code == 200 and n:
                    _THREADS_CMT["path"], _THREADS_CMT["kind"] = path, kind
                    print("  \u2705 Threads: comments endpoint found and cached.")
                    return path
            except BudgetExceeded:
                raise
            except Exception as e:
                print(f"     {path} [{kind}] -> error: {type(e).__name__}")
    print("  \u2139\uFE0F Threads: none of the tried combinations returned comments.")
    print("     404 on every line = endpoint not available;")
    print("     400/422 = endpoint exists but expects different parameters.")
    return None

def comment_url(platform, post_url, cid, it=None):
    """Permalink komentar per platform; fallback ke URL post kalau tak bisa dibentuk."""
    if not cid or not post_url:
        return post_url
    base = str(post_url).split("?")[0].rstrip("/")
    sep = "&" if "?" in str(post_url) else "?"
    p = str(platform)
    if p.startswith("Instagram"):
        return base + f"/c/{cid}/"
    if p.startswith("Facebook"):
        return str(post_url) + f"{sep}comment_id={cid}"
    if p.startswith("YouTube"):
        return str(post_url) + f"{sep}lc={cid}"
    if p.startswith("TikTok"):
        return base + f"?comment_id={cid}"
    if p.startswith("Threads") and it is not None:
        code = g(it, "code", "shortcode", "post.code")
        u = get_user(it)
        if code and u and " " not in u:
            return f"https://www.threads.net/@{u}/post/{code}"
    return post_url

def _tiktok_vid(u):
    m = re.search(r"/video/(\d+)", u or "") or re.search(r"/photo/(\d+)", u or "")
    return m.group(1) if m else ""

_CMT_PARAM = {}   # platform -> {"name":..., "val_kind":...} yang terbukti jalan (cache per sesi)

def _probe_comment_params(platform, ep, post_url, post_id):
    """Jajaki kombinasi (nama param, nilai) untuk endpoint komentar yang tampak 'tak jalan'.
    Return dict params contoh yang berhasil, atau None. Maks ~10 request, sekali per platform."""
    if platform in _CMT_PARAM:
        pr = _CMT_PARAM[platform]
        return {pr["name"]: pr["val"](post_url, post_id)} if pr else None
    clean = str(post_url).split("?")[0]
    vid = _tiktok_vid(post_url) if platform.startswith("TikTok") else ""
    vals = []
    vals.append(("url", lambda u, i: u))
    vals.append(("url", lambda u, i: str(u).split("?")[0]))
    if vid:
        for nm in ("aweme_id", "video_id", "id", "item_id", "post_id", "awemeId", "videoId"):
            vals.append((nm, (lambda v: (lambda u, i: v))(vid)))
        vals.append(("url", (lambda v: (lambda u, i: f"https://www.tiktok.com/@x/video/{v}"))(vid)))
    if post_id:
        for nm in ("id", "post_id", "media_id", "aweme_id"):
            vals.append((nm, (lambda pi: (lambda u, i: str(pi)))(post_id)))
    # buang duplikat (nama,val) untuk hemat request
    seen_pairs, uniq = set(), []
    for nm, fn in vals:
        try: key = (nm, str(fn(post_url, post_id)))
        except Exception: key = (nm, None)
        if key in seen_pairs: continue
        seen_pairs.add(key); uniq.append((nm, fn))
    vals = uniq
    print(f"  \U0001F50D {platform}: probing comment parameters (once per session, max {len(vals)} requests)...")
    for name, fn in vals:
        try:
            val = fn(post_url, post_id)
            if not val:
                continue
            r = requests.get(BASE + ep, headers=HEADERS, params={name: val}, timeout=30)
            n = len(find_items(r.json())) if r.status_code == 200 else 0
            print(f"     {name}={str(val)[:26]} -> HTTP {r.status_code}" + (f", {n} item" if r.status_code == 200 else ""))
            if r.status_code == 200 and n:
                _CMT_PARAM[platform] = {"name": name, "val": fn}
                print(f"  \u2705 {platform}: comment parameter found ('{name}') and cached.")
                return {name: val}
        except BudgetExceeded:
            raise
        except Exception as e:
            print(f"     {name} -> error: {type(e).__name__}")
    _CMT_PARAM[platform] = None
    print(f"  \u2139\uFE0F {platform}: no parameter combination returned comments.")
    try:
        val0 = vals[0][1](post_url, post_id)
        r0 = requests.get(BASE + ep, headers=HEADERS, params={vals[0][0]: val0}, timeout=30)
        body = r0.text[:300].replace("\n", " ")
        print(f"     sample response (HTTP {r0.status_code}): {body}")
        print("     \u2192 Data present but unreadable = parsing issue;")
        print("       HTTP 4xx / error message = wrong endpoint or parameter.")
    except Exception:
        pass
    return None

def fetch_comments(platform, post_url, post_id, want_raw, log, cap=50):
    ep = COMMENTS_EP.get(platform)
    if not ep and platform == "Threads":
        ep = _threads_cmt_ep(post_url, post_id)
        if not ep:
            return []
    if not ep or not post_url:
        return []
    unlimited = not cap or cap <= 0
    items, cursor, page = [], None, 0
    seen_cids = set()
    probed = False
    def _base_params():
        if platform == "Threads":
            return _threads_param(_THREADS_CMT.get("kind") or "url", post_url, post_id) \
                   or {"url": _threads_clean_url(post_url)}
        if platform in _CMT_PARAM and _CMT_PARAM[platform]:
            pr = _CMT_PARAM[platform]
            return {pr["name"]: pr["val"](post_url, post_id)}
        return {"url": post_url}
    while page < (60 if unlimited else 10):
        params = _base_params()
        if cursor:
            params[cursor[0]] = cursor[1]
        try:
            r = requests.get(BASE + ep, headers=HEADERS, params=params, timeout=45)
        except Exception:
            break
        if r.status_code != 200:
            if page == 0 and not probed and platform != "Threads" and platform not in _CMT_PARAM:
                probed = True
                if _probe_comment_params(platform, ep, post_url, post_id):
                    continue
            log(f"  {platform} comments: HTTP {r.status_code}")
            break
        data = r.json()
        batch = find_items(data)
        if not batch:
            if page == 0 and not probed and platform != "Threads" and platform not in _CMT_PARAM:
                probed = True
                if _probe_comment_params(platform, ep, post_url, post_id):
                    continue          # ulangi halaman 0 dgn parameter yang benar
            break
        fresh = []
        for b in batch:
            bu = unwrap(b)
            bk = get_id(bu) or get_text(bu)[:80]
            if bk in seen_cids:
                continue
            seen_cids.add(bk); fresh.append(b)
        if not fresh:               # halaman berisi komentar yang sama semua -> berhenti
            break
        items.extend(fresh)
        page += 1
        if not unlimited and len(items) >= cap:
            break
        cur = _find_cursor(data)
        more = data.get("has_more", data.get("more_available", data.get("has_next_page"))) if isinstance(data, dict) else None
        if not cur or (more is not None and not more):
            break
        if cursor and str(cur[1]) == str(cursor[1]):
            break
        cursor = cur
        time.sleep(0.2)
    if not unlimited:
        items = items[:cap]
    out = []
    for it_raw in items:
        it = unwrap(it_raw)
        tgl, jam = to_wib(get_ts(it))
        out.append({
            "platform": platform, "type": "comment", "parent_id": post_id,
            "date": tgl, "time": jam,
            "username": get_user(it), "user_id": get_user_id(it), "profile_image": get_profile_image(it),
            "text": get_text(it).replace("\n", " "),
            "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite"]),
            "comments": num_at(it, COMMENT_PATHS, ["reply", "comment"]),
            "shares": 0, "views": 0, "followers": get_followers(it),
            "id": get_id(it), "url": comment_url(platform, post_url, get_id(it), it), "query": "(comment)",
        })
    return out

def is_ig_post_url(u):
    ul = u.lower()
    return "instagram.com" in ul and any(k in ul for k in ["/p/", "/reel/", "/tv/"])

def _google_collect_urls(queries, is_post_url, max_items, max_pages, log, label):
    urls = []
    for q in queries:
        page = 1
        while len(urls) < max_items and page <= max_pages:
            data = google_search(q, page, log)
            if not isinstance(data, dict):
                break
            results = data.get("results") or []
            added = 0
            for res in results:
                u = (res.get("url", "") if isinstance(res, dict) else "").split("?")[0].rstrip("/")
                if u and is_post_url(u) and u not in urls:
                    urls.append(u); added += 1
            log(f"  [{label}] '{q[:44]}...' p{page}: {len(urls)} URLs")
            page += 1
            if added == 0:
                break
            time.sleep(0.4)
        if len(urls) >= max_items:
            break
    return urls

def scrape_ig_via_google(term, max_items, want_raw, log, s_ep=None, u_ep=None):
    _native = _native_keyword_search("Instagram", [
        ("/v1/instagram/search/posts", "query"), ("/v2/instagram/search", "query"),
        ("/v1/instagram/search", "keyword"), ("/v1/instagram/posts/search", "q")],
        term, max_items, want_raw, log, s_ep, u_ep)
    if _native is not None:
        return _native
    queries = [f"{term} site:instagram.com"]
    if COVERAGE["on"]:
        queries += [f'"{term}" site:instagram.com',
                    f"{term} site:instagram.com/reel",
                    f"{term} site:instagram.com/p"]
    urls = _google_collect_urls(queries, is_ig_post_url, max_items,
                                10 if COVERAGE["on"] else 8, log, "IG")
    pool = urls[:max_items]
    def _ig_fetch_chunk(chunk_urls):
        return [r for r in parallel_map(lambda u: _fetch_link(u, want_raw, log), chunk_urls, workers=6) if r]
    if not (s_ep or u_ep):
        rows = _ig_fetch_chunk(pool)
        for r in rows:
            r["query"] = term
        log(f"  Instagram(keyword): {len(rows)} posts")
        return rows
    rows, got_dated, total_inr = [], 0, 0
    for k in range(0, len(pool), 10):
        chunk = _ig_fetch_chunk(pool[k:k+10])
        chunk_inr = 0
        for row in chunk:
            row["query"] = term
            _pe = row_epoch_wib(row.get("date"), row.get("time"))
            if _pe is None:
                rows.append(row); continue
            got_dated += 1
            if (s_ep and _pe < s_ep) or (u_ep and _pe >= u_ep):
                continue
            chunk_inr += 1; total_inr += 1; rows.append(row)
        if got_dated >= 10 and chunk_inr == 0 and total_inr == 0:
            log("  Instagram: first 10+ links are all outside the time range -> stopping (saves credits)")
            break
    log(f"  Instagram(keyword): {len(rows)} posts in range")
    return rows

def _fetch_link(url, want_raw, log):
    plat = detect_platform(url)
    if plat is None:
        log(f"  {url[:48]} -> unknown platform"); return None
    if plat == "Twitter" and TWITTER_API_KEY.strip():
        return _fetch_tweet_getx(url, want_raw, log)
    path = DETAIL_EP.get(plat)
    if not path:
        log(f"  {plat} not supported by-link"); return None
    params = {"url": url}
    if plat == "YouTube":
        params["id"] = yt_id(url)
    try:
        r = requests.get(BASE + path, headers=HEADERS, params=params, timeout=45)
    except Exception as e:
        log(f"  {plat} error: {e}"); return None
    if r.status_code != 200:
        log(f"  {plat} HTTP {r.status_code}"); return None
    data = r.json()
    it = pick_post(data)
    tgl, jam = to_wib(get_ts(it))
    user = get_user(it); pid = get_id(it)
    row = {
        "platform": plat, "type": media_subtype(plat, it, url),
        "date": tgl, "time": jam, "username": user, "user_id": get_user_id(it), "profile_image": get_profile_image(it),
        "_profile": prof_ident(plat, it, user),
        "text": get_text(it).replace("\n", " "),
        "likes": num_at(it, LIKE_PATHS, ["like", "digg", "favorite", "reaction"]),
        "comments": num_at(it, COMMENT_PATHS, ["comment", "reply"]),
        "shares": num_at(it, SHARE_PATHS, ["share", "repost", "reshare"]),
        "views": num_at(it, VIEW_PATHS, ["view", "play"]),
        "followers": get_followers(it),
        "id": pid, "url": url, "query": "(by link)",
    }
    if want_raw:
        row["raw_json"] = _json.dumps(data, ensure_ascii=False)[:2000]
    return row

def fetch_retweeters(tweet_url, want_raw, log, cap=0):
    """Ambil daftar akun yang me-RETWEET sebuah tweet asli (by-link).
    Coba endpoint retweeters khusus; kalau tak ada, fallback ke search 'quoted/RT of'."""
    tid = _tweet_id(tweet_url)
    if not tid or not TWITTER_API_KEY.strip():
        return []
    n = cap if cap and cap > 0 else 200
    hdr = {"Authorization": f"Bearer {TWITTER_API_KEY.strip()}"}
    rows, seen = [], set()
    # a) endpoint retweeters khusus (beberapa nama umum)
    for path, pkey in (("/twitter/tweet/retweeters", "id"),
                       ("/twitter/tweet/retweeted_by", "id"),
                       ("/twitter/retweeters", "tweet_id")):
        try:
            r = requests.get(GETX_BASE + path, params={pkey: tid, "count": n}, headers=hdr, timeout=45)
        except BudgetExceeded:
            raise
        except Exception:
            continue
        if r.status_code != 200:
            continue
        users = find_items(r.json())
        if not users:
            continue
        # objek dari endpoint retweeters = USER (akun), bukan tweet.
        # Ambil teks tweet asli sekali untuk mengisi kolom Text retweet.
        _orig0 = _fetch_tweet_getx(tweet_url, False, lambda m: None)
        _otext = _orig0.get("text", "") if _orig0 else ""
        _oauthor = _orig0.get("username", "") if _orig0 else _handle_from_url(tweet_url)
        for u in users[:n]:
            uu = unwrap(u)
            uname = get_user(uu)
            if not uname or uname in seen:
                continue
            seen.add(uname)
            uid = get_user_id(uu)   # ini USER ID, bukan tweet id
            rt_text = f"RT @{_oauthor}: {_otext}" if _oauthor and _otext else _otext
            rows.append({
                "platform": "Twitter", "type": "retweet",
                "date": to_wib(get_ts(uu))[0], "time": to_wib(get_ts(uu))[1],
                "username": uname, "user_id": uid, "profile_image": get_profile_image(uu),
                "text": rt_text, "likes": 0, "comments": 0, "shares": 0, "views": 0,
                "followers": get_followers(uu),
                # ID unik baris = kombinasi akun+tweet (retweet native tak punya id tweet sendiri di sini)
                "id": f"{uid or uname}_{tid}", "parent_id": tid,
                # URL profil retweeter (VALID). Tidak mengarang /status/ palsu.
                "url": f"https://twitter.com/{uname}",
                "user_id_parent": get_user_id(_orig0) if _orig0 else "",
                "query": "(retweeters)",
            })
        if rows:
            log(f"  Twitter: {len(rows)} retweeter(s) via endpoint {path}")
            return _cap_retweets(rows, log)
    # b) fallback: cari retweet lewat teks 'RT @<author>:' tweet asli
    orig = _fetch_tweet_getx(tweet_url, False, lambda m: None)
    if orig and orig.get("username"):
        author = orig["username"]
        q = f'"RT @{author}:"'
        found = _tw_search_one(q, "Latest", n, want_raw, log)
        for r in found:
            r["type"] = "retweet"; r["parent_id"] = tid; r["query"] = "(retweeters/search)"
            rows.append(r)
        if rows:
            log(f"  Twitter: {len(rows)} retweet(s) via RT text search")
            return _cap_retweets(rows, log)
    log("  Twitter: retweeters are not available via endpoint or search for this tweet.")
    return []

def scrape_links(urls_all, want_raw, merge_dupes, log, progress=None):
    seq = list(urls_all)
    uniq = list(dict.fromkeys(seq))
    log(f"  fetching {len(uniq)} unique links in parallel...")
    results = parallel_map(lambda u: _fetch_link(u, want_raw, log), uniq, workers=8, progress=progress)
    cache = {u: res for u, res in zip(uniq, results)}
    order = uniq if merge_dupes else seq
    rows = [dict(cache[u]) for u in order if cache.get(u)]
    log(f"  done: {len(rows)} rows")
    return rows

# ---------------- NLP: Sentiment + Plutchik Emotion (lexicon-based, ID + EN) ----------------
# Kamus emosi Robert Plutchik (8 emosi dasar) — bilingual Indonesia + Inggris + slang + emoji.
# Dilengkapi stemming imbuhan ringan (me-/di-/ter-/ke-an/-nya, dst) dan normalisasi huruf berulang.
_EMO_WORDS = {
    "joy": [
        "senang","bahagia","gembira","suka","cinta","sayang","bangga","syukur","bersyukur","asyik","asik",
        "seru","keren","mantap","mantul","hebat","luar biasa","bagus","indah","cantik","ganteng","lucu",
        "menyenangkan","ceria","riang","puas","nikmat","enak","sukses","berhasil","menang","juara","hore",
        "alhamdulillah","selamat","terbaik","favorit","tawa","tertawa","ngakak","gemas","gemes",
        "happy","wkwk","wkwkwk","haha","hehe","hihi","yay","love","kece","cakep","top","oke banget",
        "girang","senyum","tersenyum","kebahagiaan","sukacita","menghibur","terhibur","hiburan","kocak",
        "lawak","gokil","receh","imut","menggemaskan","estetik","aesthetic","healing","adem","sejuk",
        "tenteram","tentram","lega","melegakan","syahdu","berkah","barokah","tabarakallah","meriah",
        "semarak","perayaan","merayakan","gemilang","cemerlang","berjaya","jaya","membanggakan",
        "kebanggaan","prestisius","memuaskan","terpuaskan","tercinta","kesayangan","terkasih","menyukai",
        "menikmati","jos","joss","sip","siip","oke","okee","mantap jiwa","kece badai","keren parah",
        "top markotop","worth it","worthit","recommended","gacor","pecah","best","the best","good job",
        "well done","congrats","congratulations","selamat ya","happy birthday","favoritku","seneng",
        "joy","glad","delighted","cheerful","fun","awesome","amazing","wonderful","great","excellent",
        "beautiful","lovely","enjoy","celebrate","laugh","smile","proud","blessed","grateful","win",
        "success","nice","cool","epic","lovely","adorable","fantastic","brilliant","superb","perfect",
    ],
    "trust": [
        "percaya","kepercayaan","terpercaya","tepercaya","dipercaya","mempercayai","amanah","jujur",
        "kejujuran","aman","nyaman","andal","handal","keandalan","setia","loyal","dukung","mendukung",
        "dukungan","yakin","meyakinkan","kredibel","transparan","transparansi","profesional",
        "profesionalisme","resmi","legal","sah","terjamin","garansi","jaminan","menjamin","kepastian",
        "bermanfaat","membantu","tolong","gotong royong","kerjasama","kerja sama","solid","kompak",
        "peduli","kepedulian","ramah","sopan","santun","hormat","menghormati","menghargai","apresiasi",
        "salut","terima kasih","makasih","thanks","thank you","integritas","berintegritas","akuntabel",
        "akuntabilitas","konsisten","komitmen","berkomitmen","tanggung jawab","bertanggung jawab",
        "responsif","sigap","gercep","tanggap","cekatan","kompeten","kompetensi","berpengalaman","ahli",
        "pakar","berdedikasi","dedikasi","tulus","ketulusan","ikhlas","keikhlasan","solidaritas",
        "sinergi","bersinergi","kolaborasi","berkolaborasi","koordinasi","musyawarah","mufakat",
        "harmonis","kerukunan","rukun","toleransi","toleran","melindungi","perlindungan","mengayomi",
        "pengayoman","melayani","pelayanan prima","bijak","bijaksana","arif","teladan","keteladanan",
        "panutan","terverifikasi","valid","akurat","faktual","legalitas","taat","patuh","kepatuhan",
        "disiplin","tertib","kondusif","stabil","stabilitas","aman terkendali","amanat","mengapresiasi",
        "trust","reliable","honest","support","helpful","safe","secure","credible","respect",
        "appreciate","care","loyal","faith","confidence","official","trusted","verified","accountable",
        "transparent","committed","dependable","sincere","integrity","responsive","competent",
    ],
    "fear": [
        "takut","ketakutan","ngeri","seram","serem","horor","cemas","kecemasan","khawatir","kuatir",
        "was-was","waswas","panik","gugup","grogi","gentar","ancaman","mengancam","terancam","bahaya",
        "berbahaya","rawan","risiko","resiko","darurat","gawat","teror","trauma","fobia","menakutkan",
        "menyeramkan","waspada","deg-degan","merinding","mencekam","genting","intimidasi",
        "terintimidasi","diteror","ketar-ketir","kalang kabut","paranoid","menggigil","gemetar",
        "gemetaran","mengkhawatirkan","kekhawatiran","meresahkan","keresahan","resah","gelisah",
        "kegelisahan","tidak aman","awas","hati-hati","berjaga-jaga","menghantui","menakuti","ditakuti",
        "berisiko","membahayakan","nyali ciut","bulu kuduk","insecure","overthinking","siaga",
        "fear","afraid","scared","scary","terrified","anxious","anxiety","worried","worry","panic",
        "threat","danger","dangerous","risk","emergency","horror","nightmare","dread","terrifying",
        "earthquake","quake","evacuate","evacuated","evacuation","evakuasi","dievakuasi",
        "frightening","alarming","alert",
    ],
    "surprise": [
        "kaget","terkejut","mengejutkan","kejutan","tiba-tiba","mendadak","tak terduga","tidak terduga",
        "tak disangka","gak nyangka","ga nyangka","nggak nyangka","heran","terheran","tercengang",
        "takjub","menakjubkan","wow","wah","waduh","astaga","astagfirullah","astaghfirullah",
        "masyaallah","masya allah","loh","lho","kok bisa","ternyata","viral","heboh","gempar",
        "menghebohkan","fantastis","ajaib","aneh","janggal","mengagetkan","dikejutkan","terperangah",
        "terbelalak","melongo","syok","shock","plot twist","di luar dugaan","diluar dugaan",
        "tak terbayangkan","tidak terbayangkan","fenomenal","spektakuler","wih","widih","buset",
        "busyet","gile","gilak","hah","lah kok","eh kok","speechless","terkesima","terpukau","memukau",
        "menganga","no way","seriusan","masa sih","masak sih","tumben","trending","kaget banget",
        "surprise","surprised","shocking","shocked","sudden","unexpected","unbelievable","incredible",
        "astonishing","stunning","omg","whoa","mindblowing","wtf","crazy","insane",
    ],
    "sadness": [
        "sedih","kesedihan","duka","berduka","dukacita","pilu","haru","terharu","menyedihkan",
        "memilukan","kecewa","mengecewakan","kekecewaan","menyesal","penyesalan","menangis","nangis",
        "tangis","air mata","galau","murung","suram","kelam","sepi","sunyi","kesepian","kehilangan",
        "ditinggal","putus asa","patah hati","hancur","terpuruk","nestapa","malang","kasihan",
        "prihatin","keprihatinan","turut berduka","innalillahi","rip","meninggal","wafat","tewas",
        "korban","tragedi","tragis","musibah","bencana","gagal","kegagalan","kalah","rugi","kerugian",
        "miskin","kemiskinan","derita","menderita","penderitaan","berkabung","duka mendalam",
        "belasungkawa","turut berdukacita","husnul khotimah","al-fatihah","alfatihah","miris",
        "memiriskan","ngenes","nelangsa","merana","sengsara","kesengsaraan","terlantar",
        "menelantarkan","yatim","duafa","dhuafa","memelas","iba","menyayat hati","menyayat","pedih",
        "perih","getir","muram","hampa","kecewa berat","patah semangat","putus harapan","menyerah",
        "pasrah","terpukul","berlinang","mewek","gundah","gundah gulana","lara","berpulang",
        "tutup usia","meninggal dunia","jenazah","almarhum","almarhumah","takziah","melayat",
        "kematian","gugur","korban jiwa","korban tewas","kelaparan","pengungsi","mengungsi",
        "porak-poranda","porak poranda","luluh lantak","hancur lebur","menyesakkan","berkaca-kaca",
        "mengharukan","haru biru","sedih banget","sedih bgt",
        "sad","sadness","sorrow","grief","cry","crying","tears","heartbroken","disappointed",
        "disappointing","regret","lonely","depressed","miserable","tragic","tragedy","loss","lost",
        "mourning","victim","condolences","devastated","hurt","painful","unfortunate",
    ],
    "disgust": [
        "jijik","menjijikkan","menjijikan","muak","memuakkan","mual","kotor","jorok","najis","busuk",
        "bau","amit-amit","hina","kehinaan","rendahan","murahan","norak","kampungan","alay","lebay",
        "cringe","munafik","kemunafikan","hipokrit","curang","kecurangan","licik","korup","korupsi",
        "koruptor","suap","kolusi","nepotisme","penipuan","penipu","tipu","bohong","kebohongan",
        "pembohong","hoax","hoaks","fitnah","aib","skandal","mesum","cabul","asusila","bejat","biadab",
        "keji","tercela","memalukan","ilfeel","ilfil","eneg","risih","bobrok","borok","kebusukan",
        "tercemar","sampah","pelecehan","melecehkan","menghina","penghinaan","hinaan","merendahkan",
        "nista","menistakan","penistaan","tak tahu malu","gak tau malu","tidak tahu malu","muka tebal",
        "menjilat","penjilat","cari muka","carmuk","oportunis","culas","picik","kongkalikong",
        "manipulasi","manipulatif","memanipulasi","gratifikasi","pungli","pungutan liar","penggelapan",
        "pencucian uang","penyalahgunaan","menyalahgunakan","penyelewengan","menyeleweng","selingkuh",
        "perselingkuhan","khianat","pengkhianat","pengkhianatan","berkhianat","dusta","pendusta",
        "berdusta","propaganda","pencitraan","settingan","gimmick","amoral","tidak bermoral","yuck",
        "ew","eww","hoaks murahan","bodong","abal-abal","palsu","pemalsuan",
        "gross","disgusting","disgust","nasty","filthy","sick","sickening","corrupt","corruption",
        "fraud","scam","fake","liar","lie","shameful","shameless","hypocrite","vile","repulsive",
        "pathetic","sleazy","money laundering",
    ],
    "anger": [
        "marah","kemarahan","murka","geram","kesal","kesel","sebel","sebal","jengkel","dongkol",
        "emosi","naik darah","naik pitam","benci","kebencian","membenci","dendam","amarah","berang",
        "ngamuk","mengamuk","amuk","protes","kecam","mengecam","kecaman","hujat","menghujat","hujatan",
        "caci","maki","cacian","makian","umpat","umpatan","serang","menyerang","serangan","perang",
        "konflik","bentrok","rusuh","kerusuhan","anarkis","provokasi","provokator","tolol","bodoh",
        "goblok","bego","idiot","dungu","brengsek","bangsat","sialan","kurang ajar","tidak becus",
        "gak becus","zalim","dzalim","kejam","kekejaman","brutal","sadis","anjir","anjay","edan",
        "parah","melawan","perlawanan","menentang","menolak","penolakan","protes keras","unjuk rasa",
        "demonstrasi","boikot","memboikot","mengutuk","terkutuk","laknat","melaknat","sumpah serapah",
        "memaki","mencaci","membentak","penganiayaan","menganiaya","aniaya","kekerasan","bengis",
        "ganas","beringas","membabi buta","amukan","amuk massa","ricuh","kericuhan","bentrokan",
        "tawuran","keributan","cekcok","pertikaian","bertikai","perseteruan","berseteru","permusuhan",
        "memusuhi","sakit hati","tersinggung","menyinggung","provokatif","memprovokasi","kampret",
        "bajingan","keparat","asu","jancok","jancuk","bangke","anjrit","njir","payah","lemot","lelet",
        "bertele-tele","tidak kompeten","inkompeten","amatiran","asal-asalan","ngawur","sembrono",
        "ceroboh","teledor","lalai","kelalaian","mangkrak","molor","terbengkalai","omong kosong",
        "bualan","janji palsu","ingkar janji","ingkar","php","sewot","misuh","ngegas","songong",
        "belagu","sombong","kesombongan","angkuh","arogan","arogansi","semena-mena","sewenang-wenang",
        "otoriter","represif","menindas","penindasan","tirani","zolim","menzalimi","dizalimi",
        "ketidakadilan","tidak adil","diskriminasi","diskriminatif","rasis","rasisme","intoleran",
        "intoleransi","persekusi","kriminalisasi","membungkam","gerah","jengah","muak banget",
        "angry","anger","mad","furious","rage","hate","hatred","annoyed","annoying","outrage",
        "outrageous","attack","fight","war","conflict","stupid","idiot","damn","cruel","violence",
        "violent","unfair","injustice","racist","corruptor","protest","boycott","condemn","disgrace",
        "destroy","destroying","destruction","destroyed",
    ],
    "anticipation": [
        "harap","harapan","berharap","mengharapkan","semoga","mudah-mudahan","moga","insyaallah",
        "insya allah","nantikan","menanti","menantikan","penantian","tunggu","menunggu","ditunggu",
        "tak sabar","gak sabar","ga sabar","nggak sabar","penasaran","antusias","antusiasme","rencana",
        "merencanakan","persiapan","bersiap","siap-siap","segera","akan datang","coming soon",
        "mendatang","masa depan","target","visi","misi","cita-cita","impian","mimpi","ambisi",
        "optimis","optimisme","optimistis","prediksi","proyeksi","potensi","peluang","kesempatan",
        "menjelang","countdown","launching","peluncuran","perdana","terbaru","menyongsong","menyambut",
        "digadang-gadang","diprediksi","diproyeksikan","diperkirakan","ekspektasi","menanti-nanti",
        "dinanti","dinantikan","ditunggu-tunggu","hitung mundur","sebentar lagi","tak lama lagi",
        "akan segera","bersiap-siap","penjajakan","uji coba","tahap awal","langkah awal","prospek",
        "prospektif","berpotensi","menjanjikan","harapan baru","angin segar","titik terang",
        "secercah harapan","ditargetkan","stay tuned","grand opening","pembukaan","peresmian",
        "diresmikan","groundbreaking","penasaran banget","kepo","gaskeun","gasken","ayo","ayok","yuk",
        "mari","semangat","lets go","let's go",
        "hope","hopefully","expect","expectation","await","waiting","excited","exciting","upcoming",
        "soon","future","plan","planning","optimistic","potential","opportunity","launch",
        "anticipate","looking forward","cant wait","can't wait",
    ],
}
# Kata sentimen tambahan (non-emosi spesifik)
_POS_EXTRA = [
    "baik","kebaikan","positif","untung","keuntungan","manfaat","bermutu","berkualitas","murah",
    "hemat","gratis","promo","diskon","lancar","mudah","praktis","efektif","efisien","bersih","sehat",
    "kuat","maju","kemajuan","berkembang","meningkat","peningkatan","prestasi","penghargaan","inovasi",
    "inovatif","solusi","damai","perdamaian","sejahtera","kesejahteraan","adil","keadilan","makmur",
    "rekomendasi","membaik","perbaikan","pemulihan","pulih","tumbuh","pertumbuhan","surplus",
    "efektivitas","efisiensi","produktif","produktivitas","optimal","maksimal","unggul","keunggulan",
    "berprestasi","terakreditasi","bermartabat","asri","tertata","rapi","istimewa","spesial",
    "memadai","layak","terjangkau","ekonomis","bergizi","higienis","ramah lingkungan","berkelanjutan",
    "inklusif","merata","pemerataan","swasembada","mandiri","kemandirian","berdaya","pemberdayaan",
    "good","nice","best","better","improve","improvement","clean","healthy","strong","progress",
    "peace","fair","prosperous","benefit","quality","affordable","easy","effective","solution",
    "innovative","sustainable","inclusive",
]
_NEG_EXTRA = [
    "buruk","jelek","negatif","rusak","kerusakan","merusak","masalah","bermasalah","gangguan",
    "terganggu","sulit","susah","ribet","rumit","macet","kemacetan","banjir","gempa","kebakaran",
    "kecelakaan","luka","sakit","penyakit","wabah","virus","kriminal","kejahatan","pencurian",
    "maling","curi","rampok","begal","pembunuhan","bunuh","narkoba","narkotika","mahal","kemahalan",
    "langka","kelangkaan","krisis","inflasi","utang","hutang","pengangguran","phk","gagal bayar",
    "bangkrut","pailit","melanggar","pelanggaran","ilegal","sengketa","polemik","kontroversi",
    "kontroversial","tersangka","terdakwa","vonis","penjara","memburuk","penurunan","anjlok",
    "merosot","kemerosotan","defisit","boros","pemborosan","mubazir","sia-sia","percuma",
    "gagal total","berantakan","kacau","kekacauan","amburadul","semrawut","kumuh","polusi",
    "pencemaran","beracun","pandemi","parah banget","memprihatinkan","rusak parah","ambruk","roboh",
    "longsor","tanah longsor","gempa bumi","tsunami","erupsi","kebanjiran","kekeringan","gagal panen",
    "naik harga","kenaikan harga","membebani","terbebani","kredit macet","phk massal","dirumahkan",
    "menganggur","kemiskinan ekstrem","stunting","gizi buruk","putus sekolah","kriminalitas",
    "curanmor","jambret","penculikan","penipuan online","judi online","judol","pinjol ilegal",
    "pinjol","pidana","gugatan","digugat","dipenjara","ditahan","penahanan","buron","sanksi","denda",
    "error","eror","gangguan sistem","server down","mati listrik","mati lampu","pemadaman",
    "trouble","komplain","keluhan","mengeluh",
    "keluh kesah","aduan","zonk","bapuk",
    "bad","worse","worst","broken","problem","difficult","hard","expensive","crime","criminal",
    "illegal","crisis","debt","unemployed","fail","failure","accident","injury","disease","virus",
    "controversy","corruptions","lawsuit","scandalous","polluted","toxic",
]
_EMOJI_EMO = {
    "😀":"joy","😃":"joy","😄":"joy","😁":"joy","😆":"joy","😊":"joy","🙂":"joy","😍":"joy","🥰":"joy",
    "😘":"joy","😂":"joy","🤣":"joy","😹":"joy","☺":"joy","😌":"joy","🤩":"joy","😻":"joy","🥳":"joy",
    "❤":"joy","🧡":"joy","💛":"joy","💚":"joy","💙":"joy","💜":"joy","💖":"joy","💕":"joy","💗":"joy",
    "💓":"joy","💞":"joy","💘":"joy","♥":"joy","🙌":"joy","👏":"joy","🔥":"joy","💯":"joy","✨":"joy",
    "🎉":"joy","🎊":"joy","😅":"joy",
    "👍":"trust","🤝":"trust","🙏":"trust","✅":"trust","☑":"trust",
    "💪":"anticipation","👀":"anticipation","🤞":"anticipation","⏳":"anticipation","🚀":"anticipation",
    "⌛":"anticipation",
    "😢":"sadness","😭":"sadness","😔":"sadness","😞":"sadness","😥":"sadness","💔":"sadness",
    "🥀":"sadness","😿":"sadness","🙁":"sadness","☹":"sadness","😩":"sadness","😫":"sadness",
    "😪":"sadness","🥺":"sadness",
    "😡":"anger","🤬":"anger","😠":"anger","👿":"anger","💢":"anger","🖕":"anger","😤":"anger",
    "😱":"fear","😨":"fear","😰":"fear","😖":"fear","😬":"fear",
    "😲":"surprise","😮":"surprise","😯":"surprise","🤯":"surprise","😳":"surprise","‼":"surprise",
    "⁉":"surprise",
    "🤢":"disgust","🤮":"disgust","💩":"disgust","🙄":"disgust","😒":"disgust","🤥":"disgust",
    "👎":"disgust",
}
_NEGATORS = {"tidak","tak","tdk","gak","ga","nggak","ngga","engga","enggak","bukan","tanpa","belum",
             "jangan","anti","no","not","never","dont","don't","isnt","isn't","wont","won't","didnt",
             "didn't","cant","can't","doesnt","doesn't","aint","ain't","kurang"}
_NEG_BRIDGE = {"terlalu","begitu","cukup","akan","bisa","sangat","semakin","lagi","juga","pernah",
               "usah","banget","terlihat","tampak","merasa","perlu","mau","ada","begini","gitu","so",
               "very","really","that","too"}
_EMO_OPP = {"joy":"sadness","sadness":"joy","trust":"disgust","disgust":"trust",
            "fear":"anger","anger":"fear","surprise":"anticipation","anticipation":"surprise"}
_EMO_POL = {"joy":1,"trust":1,"anticipation":1,"surprise":0,"sadness":-1,"disgust":-1,"anger":-1,"fear":-1}

def _build_lex():
    uni, phr = {}, {}
    def put(w, tag):
        w = w.lower().strip()
        if " " in w or "-" in w:
            phr[w.replace("-", " ")] = tag
        else:
            uni[w] = tag
    for emo, words in _EMO_WORDS.items():
        for w in words:
            put(w, ("E", emo))
    for w in _POS_EXTRA:
        put(w, ("S", 1))
    for w in _NEG_EXTRA:
        put(w, ("S", -1))
    return uni, phr
_LEX_UNI, _LEX_PHR = _build_lex()
_TOK_RE = re.compile(r"[a-zA-Z\u00C0-\u024F']+")
_ELONG3 = re.compile(r"(.)\1{2,}")

_STEM_PREFIXES = (("meng", ("", "k", "g")), ("meny", ("s",)), ("mem", ("", "p", "b")),
                  ("men", ("", "t")), ("peng", ("", "k")), ("peny", ("s",)), ("pem", ("", "p")),
                  ("pen", ("", "t")), ("ber", ("",)), ("ter", ("",)), ("per", ("",)),
                  ("me", ("",)), ("di", ("",)), ("ke", ("",)), ("se", ("",)), ("pe", ("",)))
_STEM_SUFFIXES = ("lah", "kah", "tah", "pun", "nya", "ku", "mu", "kan", "an", "i")

# Topic/institution nouns whose stem is emotional but the word itself is neutral
# (e.g. "kebijakan" = policy, not "bijak" = wise; "pemilu" = election, not "pilu" = sad).
_NEUTRAL_WORDS = {
    "kebijakan", "kebijakannya", "pemilu", "pemilihan", "kesehatan", "keamanan", "keselamatan",
    "ketertiban", "kesejahteraan", "perlindungan", "terbaru", "keadilan", "kebersihan", "pembangunan",
    "pendidikan", "kemanusiaan", "keuangan", "kepercayaan", "kesenian", "perdamaian",
}
_NEUTRAL_PHRASES = ("rumah sakit", "sakit jiwa", "kementerian kesehatan", "dinas kesehatan",
                    "badan keamanan", "lembaga keamanan", "hari raya", "selamat pagi",
                    "selamat siang", "selamat sore", "selamat malam")

def _lex_lookup(w):
    """Look a word up: exact form -> collapsed repeated letters -> strip Indonesian affixes."""
    if w in _NEUTRAL_WORDS:
        return None
    hit = _LEX_UNI.get(w)
    if hit:
        return hit
    cands, seen = [], {w}
    for x in (_ELONG3.sub(r"\1\1", w), _ELONG3.sub(r"\1", w)):
        if x not in seen:
            seen.add(x); cands.append(x)
    bases = [w] + cands
    forms = list(bases)
    for b in bases:
        cur = {b}
        for _ in range(2):
            nxt = set()
            for f in cur:
                for suf in _STEM_SUFFIXES:
                    if f.endswith(suf) and len(f) - len(suf) >= 3:
                        s = f[:-len(suf)]
                        if s not in seen:
                            seen.add(s); nxt.add(s); forms.append(s)
            cur = nxt or cur
    for f in list(forms):
        for pre, reps in _STEM_PREFIXES:
            if f.startswith(pre) and len(f) - len(pre) >= 3:
                rest = f[len(pre):]
                for rp in reps:
                    x = rp + rest
                    if x not in seen:
                        seen.add(x); forms.append(x)
    for f in forms:
        hit = _LEX_UNI.get(f)
        if hit:
            return hit
    return None

def _analyze_lex(text):
    """Analisis leksikon: return (sentiment, emotion, emo_score dict)."""
    if not text or not isinstance(text, str):
        return "neutral", "neutral", {e: 0.0 for e in _EMO_WORDS}
    raw = text
    t = text.lower()
    t = re.sub(r"https?://\S+|www\.\S+", " ", t)   # buang URL
    t = re.sub(r"@\w+", " ", t)                     # buang mention
    t = t.replace("#", " ")
    for ph in _NEUTRAL_PHRASES:
        if ph in t:
            t = t.replace(ph, " ")
    emo_score = {e: 0.0 for e in _EMO_WORDS}
    sent = 0.0
    # 1) emoji
    for ch, emo in _EMOJI_EMO.items():
        n = raw.count(ch)
        if n:
            emo_score[emo] += n
            sent += _EMO_POL[emo] * n
    # 2) frasa (multi-kata)
    for ph, tag in _LEX_PHR.items():
        n = t.count(ph)
        if n:
            if tag[0] == "E":
                emo_score[tag[1]] += n; sent += _EMO_POL[tag[1]] * n
            else:
                sent += tag[1] * n
    # 3) kata tunggal + stemming + negasi (jarak 1-2)
    toks = _TOK_RE.findall(t)
    for i, w in enumerate(toks):
        if w in _NEGATORS or w in _NEG_BRIDGE:
            continue
        tag = _lex_lookup(w)
        if not tag:
            continue
        neg = (i > 0 and toks[i-1] in _NEGATORS) or \
              (i > 1 and toks[i-2] in _NEGATORS and toks[i-1] in _NEG_BRIDGE)
        if tag[0] == "E":
            emo = _EMO_OPP[tag[1]] if neg else tag[1]
            emo_score[emo] += 1
            sent += _EMO_POL[emo]
        else:
            sent += -tag[1] if neg else tag[1]
    best = max(emo_score, key=emo_score.get)
    emotion = best if emo_score[best] > 0 else "neutral"
    if sent > 0:
        sentiment = "positive"
    elif sent < 0:
        sentiment = "negative"
    elif emotion != "neutral" and _EMO_POL[emotion] > 0:
        sentiment = "positive"
    elif emotion != "neutral" and _EMO_POL[emotion] < 0:
        sentiment = "negative"
    else:
        sentiment = "neutral"
    return sentiment, emotion, emo_score

def analyze_text(text):
    s, e, _ = _analyze_lex(text)
    return s, e

# ---------------- Mesin AI (model bahasa) — hybrid dengan leksikon ----------------
NLP_MODE = {"mode": "ai"}   # "ai" | "lexicon"
_NLP = {"tried": False, "ok": False}
_EMOT_MAP = {"happy": "joy", "love": "trust", "sadness": "sadness", "anger": "anger", "fear": "fear"}
_EN_EMO_MAP = {"joy": "joy", "surprise": "surprise", "sadness": "sadness", "anger": "anger",
               "fear": "fear", "disgust": "disgust", "neutral": "neutral"}
_ID_HINT = {"yang","dan","di","ini","itu","tidak","ada","dengan","untuk","dari","ke","pada","saya",
            "kamu","kita","akan","sudah","belum","bisa","gak","nggak","aja","banget","kalau","karena",
            "atau","juga","mau","udah","lagi","sama","kok","sih","dong","deh","nya","pak","bu","itu"}
_EN_HINT = {"the","and","is","are","of","to","in","that","for","it","this","was","with","on","you",
            "have","not","be","at","as","by","from"}

def _lang_is_id(text):
    toks = set(re.findall(r"[a-z]+", (text or "").lower()))
    idn = len(toks & _ID_HINT); eng = len(toks & _EN_HINT)
    return idn >= eng

def _clean_for_model(t):
    t = re.sub(r"https?://\S+|www\.\S+", " ", t or "")
    return re.sub(r"\s+", " ", t).strip()[:600]

def _load_ai_models():
    if _NLP["tried"]:
        return _NLP["ok"]
    _NLP["tried"] = True
    try:
        try:
            from transformers import pipeline
            import torch
        except ImportError:
            print("  AI analysis needs extra packages that are not installed (run install-ai-windows.bat or")
            print("  install-ai-mac-linux.sh once). Using the lexicon for this run.")
            _NLP["tried"] = False   # check again on the next run, after the user installs them
            return False
        dev = 0 if torch.cuda.is_available() else -1
        _NLP["gpu"] = (dev == 0)
        print("  Loading AI models (first time downloads ~2 GB and takes a few minutes"
              + (", using GPU" if dev == 0 else ", using CPU") + ")...")
        _NLP["sent"] = pipeline("sentiment-analysis",
                                model="cardiffnlp/twitter-xlm-roberta-base-sentiment",
                                device=dev, truncation=True, max_length=128)
        _NLP["emo_id"] = pipeline("text-classification",
                                  model="StevenLimcorn/indonesian-roberta-base-emotion-classifier",
                                  device=dev, truncation=True, max_length=128)
        _NLP["emo_en"] = pipeline("text-classification",
                                  model="j-hartmann/emotion-english-distilroberta-base",
                                  device=dev, truncation=True, max_length=128)
        _NLP["ok"] = True
        print("  \u2705 AI models ready.")
    except Exception as e:
        print(f"  AI models failed to load ({type(e).__name__}: {str(e)[:70]}) \u2014 using the lexicon instead.")
        _NLP["ok"] = False
        _NLP["tried"] = False   # retry on the next run (e.g. after the internet is back)
    return _NLP["ok"]

def _merge_emotion(model_emo, model_conf, lex_emo, lex_scores):
    # leksikon menang untuk kelas yang model tidak kenal, bila sinyalnya kuat
    if lex_emo in ("anticipation", "surprise", "disgust") and lex_scores.get(lex_emo, 0) >= 2:
        return lex_emo
    if model_emo != "neutral" and model_conf >= 0.50:
        return model_emo
    if lex_emo != "neutral":
        return lex_emo
    if model_emo != "neutral" and model_conf >= 0.35:
        return model_emo
    return "neutral"

def analyze_batch(texts):
    """Analisis massal: mode AI (hybrid model+leksikon) atau leksikon murni. Return list (sent, emo)."""
    texts = [t if isinstance(t, str) else "" for t in texts]
    lex = [_analyze_lex(t) for t in texts]
    if NLP_MODE["mode"] != "ai" or not texts or not _load_ai_models():
        return [(s, e) for s, e, _ in lex]
    try:
        clean = [_clean_for_model(t) or "-" for t in texts]
        if len(clean) > 300 and not _NLP.get("gpu"):
            print(f"  \u23F3 {len(clean)} rows on CPU \u2014 this can take a few minutes. Choose Lexicon for instant results.")
        sents = []
        for k in range(0, len(clean), 256):
            sents.extend(_NLP["sent"](clean[k:k+256], batch_size=32))
            if len(clean) > 256:
                print(f"  \U0001F9E0 Sentiment analysis: {min(k+256, len(clean))}/{len(clean)} rows...")
        idx_id = [i for i, t in enumerate(clean) if _lang_is_id(t)]
        idx_en = [i for i in range(len(clean)) if i not in idx_id]
        emo_out = [("neutral", 0.0)] * len(clean)
        if idx_id:
            for i, r in zip(idx_id, _NLP["emo_id"]([clean[i] for i in idx_id], batch_size=32)):
                emo_out[i] = (_EMOT_MAP.get(str(r["label"]).lower(), "neutral"), float(r["score"]))
        if idx_en:
            for i, r in zip(idx_en, _NLP["emo_en"]([clean[i] for i in idx_en], batch_size=32)):
                emo_out[i] = (_EN_EMO_MAP.get(str(r["label"]).lower(), "neutral"), float(r["score"]))
        out = []
        for i in range(len(texts)):
            if clean[i] == "-":
                out.append(("neutral", "neutral"))   # empty / link-only text: nothing to analyse
                continue
            slabel = str(sents[i]["label"]).lower()
            sent = slabel if slabel in ("positive", "negative", "neutral") else lex[i][0]
            emo = _merge_emotion(emo_out[i][0], emo_out[i][1], lex[i][1], lex[i][2])
            out.append((sent, emo))
        return out
    except Exception as e:
        print(f"  \u26A0\uFE0F AI inference failed ({type(e).__name__}) \u2014 these rows use the lexicon.")
        return [(s, e) for s, e, _ in lex]

# ---------------- Layout output mengikuti Sample_Crawling.xlsx ----------------
SAMPLE_ORDER = ["No","Sentiment","Username","User ID","Follower","Text","Date","Time","Likes","Reply",
                "Shares","Url","Emotion","User Image","Platform","Sub Type","Views Count","Gender",
                "Age Category","Content Location","Religion","Ethnic","Phsychography","Parent ID"]
RENAME_MAP = {"username":"Username","user_id":"User ID","followers":"Follower","text":"Text","date":"Date",
              "time":"Time","likes":"Likes","comments":"Reply","shares":"Shares","url":"Url",
              "profile_image":"User Image","platform":"Platform","type":"Sub Type","views":"Views Count",
              "parent_id":"Parent ID","id":"ID","user_id_parent":"User ID Parent",
              "query":"Query","raw_json":"Raw JSON"}
MANUAL_COLS = ["Gender","Age Category","Content Location","Religion","Ethnic","Phsychography"]

def add_rekap_codebook(xp, df):
    """Tambahkan sheet Rekap (statistik siap laporan) + Codebook ke file Excel. Lokal, tanpa kredit."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = load_workbook(xp)
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    sec_font = Font(bold=True, size=12, color="1F4E79")

    # ---------- Sheet Rekap ----------
    ws = wb.create_sheet("Summary")
    r = 1
    def sec(title):
        nonlocal r
        ws.cell(row=r, column=1, value=title).font = sec_font; r += 1
    def table(headers, rows):
        nonlocal r
        for j, h in enumerate(headers, 1):
            c = ws.cell(row=r, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill
        r += 1
        for row in rows:
            for j, v in enumerate(row, 1):
                ws.cell(row=r, column=j, value=v)
            r += 1
        r += 1

    posts = df[df["Sub Type"] != "comment"] if "Sub Type" in df.columns else df
    cmts = df[df["Sub Type"] == "comment"] if "Sub Type" in df.columns else df.iloc[0:0]
    sec("OVERVIEW")
    dr = ""
    if "Date" in df.columns and df["Date"].notna().any():
        ds = df["Date"].astype(str)
        ds = ds[ds.str.match(r"\d{4}-\d{2}-\d{2}")]
        if len(ds):
            dr = f"{ds.min()} to {ds.max()}"
    table(["Metric", "Value"], [
        ["Total rows", len(df)],
        ["Posts (post/tweet/retweet/quoted)", len(posts)],
        ["Comments/replies", len(cmts)],
        ["Unique accounts", df["Username"].nunique() if "Username" in df.columns else ""],
        ["Date range", dr],
    ])

    if "Platform" in df.columns:
        sec("BY PLATFORM")
        vc = df.groupby("Platform").agg(
            Rows=("Platform", "size"),
            Likes=("Likes", lambda s: int(__import__("pandas").to_numeric(s, errors="coerce").fillna(0).sum())),
            Reply=("Reply", lambda s: int(__import__("pandas").to_numeric(s, errors="coerce").fillna(0).sum())),
            Shares=("Shares", lambda s: int(__import__("pandas").to_numeric(s, errors="coerce").fillna(0).sum())),
        )
        table(["Platform", "Rows", "Total Likes", "Total Reply", "Total Shares"],
              [[i] + list(row) for i, row in vc.iterrows()])

    for col, title in (("Sentiment", "SENTIMENT DISTRIBUTION"), ("Emotion", "EMOTION DISTRIBUTION (PLUTCHIK)")):
        if col in df.columns:
            sec(title)
            vc = df[col].value_counts()
            tot = max(1, int(vc.sum()))
            table([col, "Count", "%"], [[k, int(v), f"{100*int(v)/tot:.1f}%"] for k, v in vc.items()])

    if {"Username", "Likes"} <= set(df.columns) and len(posts):
        sec("TOP 10 MOST ACTIVE ACCOUNTS (posts)")
        vc = posts["Username"].value_counts().head(10)
        table(["Username", "Posts"], [[k, int(v)] for k, v in vc.items()])
        import pandas as _pd
        eng = posts.assign(_e=_pd.to_numeric(posts["Likes"], errors="coerce").fillna(0)
                              + _pd.to_numeric(posts.get("Reply"), errors="coerce").fillna(0)
                              + _pd.to_numeric(posts.get("Shares"), errors="coerce").fillna(0))
        top = eng.sort_values("_e", ascending=False).head(10)
        sec("TOP 10 POSTS BY ENGAGEMENT (likes+reply+shares)")
        table(["Username", "Engagement", "Text (start)"],
              [[t["Username"], int(t["_e"]), str(t.get("Text", ""))[:80]] for _, t in top.iterrows()])

    if "Date" in df.columns:
        ds = df["Date"].astype(str)
        ok = ds.str.match(r"\d{4}-\d{2}-\d{2}")
        if ok.any():
            sec("VOLUME BY DATE")
            vc = ds[ok].value_counts().sort_index()
            table(["Date", "Rows"], [[k, int(v)] for k, v in vc.items()])

    for col_cells in ws.columns:
        w = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(60, max(12, w + 2))

    # ---------- Sheet Codebook ----------
    cb = wb.create_sheet("Codebook")
    cb_rows = [
        ("No", "Row number"),
        ("Sentiment", "Automatic sentiment (AI model and/or ID+EN+emoji lexicon): positive / negative / neutral"),
        ("Username", "Account handle (no spaces). Facebook commenters may show a display name"),
        ("User ID", "Unique account ID on the platform (stable even if the username changes)"),
        ("Follower", "Account follower count (when available / enabled)"),
        ("Text", "Post/comment text. Truncated retweets are rebuilt from the original tweet"),
        ("Date / Time", "Publish date & time, WIB (GMT+7)"),
        ("Likes / Reply / Shares", "Engagement metrics. Retweets use the original tweet's metrics when their own are empty"),
        ("Url", "Direct link to the post"),
        ("Emotion", "Dominant Plutchik emotion: joy, trust, fear, surprise, sadness, disgust, anger, anticipation, neutral"),
        ("User Image", "Profile picture URL"),
        ("Platform", "twitter / tiktok / youtube / instagram / facebook / threads"),
        ("Sub Type", "Twitter: tweet/retweet/quoted/reply \u2022 Instagram: post/image/video/carousel/comment \u2022 "
                      "Facebook: post/video/link/comment \u2022 TikTok/YouTube/Threads: post/comment"),
        ("Views Count", "View count (when the platform provides it)"),
        ("Gender to Phsychography", "Manual coding columns — intentionally empty, filled in by the researcher"),
        ("Parent ID", "For comments: ID of the parent post"),
        ("ID", "Unique ID of this post/comment"),
        ("User ID Parent", "User ID of the parent post's author (posts: own user id) — for conversation clustering"),
        ("Query", "Row source: search keyword, (by link), profile:..., (comment), (reply)"),
    ]
    c1 = cb.cell(row=1, column=1, value="Column"); c2 = cb.cell(row=1, column=2, value="Description")
    for c in (c1, c2): c.font = hdr_font; c.fill = hdr_fill
    for i, (k, v) in enumerate(cb_rows, 2):
        cb.cell(row=i, column=1, value=k).font = Font(bold=True)
        cell = cb.cell(row=i, column=2, value=v)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    cb.column_dimensions["A"].width = 26
    cb.column_dimensions["B"].width = 95
    wb.save(xp)

def to_sample_layout(df):
    df = df.copy()
    texts = df["text"].fillna("") if "text" in df.columns else pd.Series([""] * len(df))
    res = analyze_batch(list(texts))
    df["Sentiment"] = [r[0] for r in res]
    df["Emotion"] = [r[1] for r in res]
    if "platform" in df.columns:
        df["platform"] = df["platform"].astype(str).str.lower()
    if "type" in df.columns and "platform" in df.columns:
        tw = df["platform"] == "twitter"
        df.loc[tw & (df["type"] == "post"), "type"] = "tweet"
        df.loc[tw & (df["type"] == "comment"), "type"] = "reply"
        if "text" in df.columns:
            rt = df["text"].fillna("").astype(str).str.startswith("RT @")
            df.loc[tw & rt & (df["type"] == "tweet"), "type"] = "retweet"
    if "username" in df.columns and "url" in df.columns:
        def _fix_user(row):
            u = str(row.get("username") or "")
            if row.get("type") == "comment":
                return u          # url komentar = url post induk; jangan dipakai menimpa nama komentator
            if u and " " not in u:
                return u
            h = _handle_from_url(str(row.get("url") or ""))
            return h or u
        df["username"] = df.apply(_fix_user, axis=1)
    df = df.rename(columns={k: v for k, v in RENAME_MAP.items() if k in df.columns})
    for c in MANUAL_COLS:
        if c not in df.columns:
            df[c] = None
    if "Sub Type" not in df.columns:
        df["Sub Type"] = "post"
    # Every output field is always present (empty when the platform does not provide it).
    for c in SAMPLE_ORDER + ["ID", "User ID Parent", "Query"]:
        if c != "No" and c not in df.columns:
            df[c] = ""
    df.insert(0, "No", range(1, len(df) + 1))
    front = [c for c in SAMPLE_ORDER if c in df.columns]
    _pref = ["ID", "User ID Parent", "Query", "Raw JSON"]
    extras = [c for c in _pref if c in df.columns and c not in front]
    extras += [c for c in df.columns if c not in front and c not in extras]
    return df[front + extras]

