"""
新刊リリース情報を取得し、Googleスプレッドシート・カレンダー・Discordに反映するスクリプト。
GAS版からの移植版（HTML解析をBeautifulSoup+検証済みロジックに置き換え、6分の実行時間制限を撤廃）。

必要な環境変数：
  GOOGLE_SERVICE_ACCOUNT_JSON  サービスアカウントのJSONキー（文字列そのまま）
  SPREADSHEET_ID               対象スプレッドシートのID
  SHEET_NAME                   対象シート名（省略時 "シート1"）
  CALENDAR_ID                  Googleカレンダーの CalendarID
  DISCORD_WEBHOOK_URL          Discord Webhook URL（任意。未設定なら通知はスキップ）
  DAILY_LIMIT                  1日あたりの自動チェック件数上限（任意。省略時40）

シート列構成：
  A: 作品タイトル   B: 最新巻数   C: 発売日   D: ジャンル
  E: ステータス（最終処理結果の表示用。処理対象の判定には使わない）
  F: 最終更新日時（この日時が古い行から優先的にチェックされる）
  G: Discord通知フラグ
  H: 手動確認フラグ（自動検索で解決できない作品の印。TRUEの間は自動検索を
     スキップしF列だけ更新する。月次まとめDiscord通知の対象にもなる）
  I: 手動確認の検索リンク（シート側の数式。スクリプトは読み書きしない）
  S1: （内部管理用）本日発売ダイジェストの送信済み日付
  J: 登録反映チェック（B・C列を手動で埋めてTRUEにするとカレンダー登録・
     Discord通知を行い、処理後に自動でFALSEへ戻る）
  K: 基準の書名  L: ISBN  M: レーベル  N: 著者
     （前回確認できた本。次回の判定の基準にする。巻数がB列と違う場合は使わず推測し直す。
       K列を空にすると推測し直す）
  O: 要確認（自動では記録せず、人が確認する内容。解決すると空になる）
"""

import json
import os
import re
import time
import unicodedata
from datetime import datetime, timedelta, timezone

# GitHub Actionsのランナーは標準でUTCで動作するため、datetime.now()を
# そのまま使うとJST（日本時間）の午前0時〜9時台の実行時に日付が1日
# 古く記録されてしまう（例: JST 8/11 6:00 = UTC 8/10 21:00）。
# F列・C列に記録する「今日」の日付は、常にJST基準で統一する。
JST = timezone(timedelta(hours=9))


def now_jst():
    return datetime.now(JST)

import gspread
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ------------------------------------------------------------
# 設定
# ------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar",
]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SHEET_NAME = os.environ.get("SHEET_NAME", "シート1")
CALENDAR_ID = os.environ["CALENDAR_ID"]
DISCORD_WEBHOOK_URL = os.environ.get("DISCORD_WEBHOOK_URL", "")
SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

GENRE_URLS = {
    "ライトノベル": "https://novel.sumikko.info/search/?k=",
    "文庫": "https://bunko.sumikko.info/search/?k=",
    "コミック": "https://comic.sumikko.info/search/?k=",
}

REQUEST_INTERVAL_SEC = 1.0  # サイトへの配慮（リクエスト間隔）

# 新刊は数カ月おきにしか出ないため、毎日全件チェックする必要はない。
# 1日あたりのチェック件数に上限を設け、最終更新日時が古い行から
# 優先的に処理することで、リクエスト数・実行時間・APIクォータを節約する。
# 例：200件登録時、DAILY_LIMIT=40なら約5日で全件が一巡する。
DAILY_LIMIT = int(os.environ.get("DAILY_LIMIT") or "40")

# GitHub ActionsのランナーIPが一時的に不安定なことがあるため、
# 接続失敗時はリトライする（診断の結果、requestsライブラリ自体やUser-Agentは
# 原因ではなく、実行のたびに変わるランナーIPの一時的な問題と判明したため）
MAX_RETRIES = 3
RETRY_BACKOFF_SEC = 5

# 接続エラーがこの回数連続で発生したら、それ以上粘らずに処理を中断する。
# GitHub Actionsの同一ジョブ内ではランナーのIPが変わらないため、
# 個々のリクエストをリトライし続けても無駄になる。中断してジョブ自体を
# 失敗させ、ワークフロー側で新しいランナー（＝新しいIP）でやり直す。
MAX_CONSECUTIVE_NETWORK_FAILURES = 3


class TooManyNetworkFailures(Exception):
    """接続エラーが連続で発生し、処理を中断したことを示す例外。"""
    pass

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


# ------------------------------------------------------------
# 認証・APIクライアント
# ------------------------------------------------------------
def get_credentials():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    return Credentials.from_service_account_info(info, scopes=SCOPES)


# ------------------------------------------------------------
# タイトル・検索関連ユーティリティ
# ------------------------------------------------------------

# 波ダッシュ「〜」(U+301C)、全角チルダ「～」(U+FF5E)、
# スウングダッシュ「⁓」(U+2053) はすべて見た目が似ているが別の文字コードで、
# NFKC正規化では統一されない（全角チルダのみASCIIの「~」に変換され、
# 波ダッシュは変換対象外）。サブタイトルを「〜」で囲む作品名が非常に多く、
# サイト側の表記とスプレッドシート側の表記でこれらの文字が食い違うと
# 完全一致判定に失敗するため、比較前にすべて同じ文字へ統一する。
TILDE_VARIANTS_PATTERN = re.compile(r"[\u301C\uFF5E\u2053~]")


def normalize_title(s):
    """タイトル比較用の正規化（全角/半角統一・波ダッシュ類統一・空白除去・小文字化）"""
    s = unicodedata.normalize("NFKC", s)
    s = TILDE_VARIANTS_PATTERN.sub("~", s)
    s = re.sub(r"[\s\u3000]+", "", s)
    return s.lower()


def simplify_title_for_search(title):
    """検索クエリ用にタイトルから記号類を除去"""
    title = re.sub(r"[<>{}\[\]【】「」『』（）()!！〜～・、。:：;]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title


DATE_PATTERN = re.compile(r"([0-9]{2})年([0-9]{1,2})月([0-9]{1,2})日\(([月火水木金土日])\)")

# ジャンル（スプレッドシートD列）ごとに、type-tag（カテゴリ表示）の
# 許可リストを切り替える（ホワイトリスト方式）。
# サイトが検索対象のジャンルによって別ドメイン（comic./novel./bunko.）に
# 分かれているのに加え、同じサイト内でもtype-tagの値がジャンルによって
# 異なるため、ジャンルごとに許可するtype-tagを個別に定義する。
#
# 実データで確認済みの値：
#   コミック     : "コミック"
#   ライトノベル : "ライトノベル" / "TL" / "ジュニアノベル"
#   文庫         : 未確認（"文庫"を仮設定。誤除外が出たら調整）
ALLOWED_CATEGORIES = {
    "コミック": {"コミック", "単行本"},
    "ライトノベル": {"ライトノベル", "TL", "ジュニアノベル"},
    "文庫": {"文庫"},
}

# 特装版・関連本と外伝は、本編の新刊として扱わない（作品名より後ろの部分で判定する）。
NG_PATTERN = re.compile(
    r"特装版|限定版|同梱版|ガイドブック|キャラクターブック|ファンブック|スペシャルブック"
    r"|BOOK|MEMORIAL|祝ノ書|画集|Art ?Works|シナリオ集|ぬり絵|カレンダー|巻セット|完結セット|セット\(全",
    re.I,
)
SPIN_PATTERN = re.compile(r"外伝|番外編|短編集|スピンオフ|蛇足編|-(another|extra|after|before)-", re.I)
# 巻数（15 / 12.5 / (5) / XIII）と、その直前の言葉（本編の目印）を切り出す。
# "3年生編5" の "3" のように「年」が続く数字は巻数とみなさない。
TOKEN_PATTERN = re.compile(
    r"^(.*?)(\(?\d+(?:\.\d+)?\)?|(?<![A-Za-z])[IVXL]+(?![A-Za-z]))(?=$|\s|\(|[^\d.A-Za-z年])"
)


def split_after_title(name, title):
    """書名が作品名で始まれば残りを返す。始まらなければ None。"""
    target = normalize_title(title)
    n = unicodedata.normalize("NFKC", name)
    for i in range(len(n) + 1):
        head = normalize_title(n[:i])
        if head == target:
            return TILDE_VARIANTS_PATTERN.sub("~", n[i:]).strip()
        if not target.startswith(head):
            return None
    return None


def token_value(token):
    t = token.strip("()")
    if t[0].isdigit():
        return float(t)
    total = prev = 0
    for ch in reversed(t):
        v = {"I": 1, "V": 5, "X": 10, "L": 50}[ch]
        total += -v if v < prev else v
        prev = max(prev, v)
    return float(total)


def signature_and_volume(rest):
    """作品名より後ろの部分から (本編の目印, 巻数) を返す。巻数がなければ (None, None)。"""
    m = TOKEN_PATTERN.match(rest)
    if not m:
        return None, None
    value = token_value(m.group(2))
    # ローマ数字の XV がサイトで "105"（10 と 5）と表記される。末尾の (15) と一致するときだけ直す。
    split = re.fullmatch(r"([1-4]0)(\d)", m.group(2))
    tail = re.search(r"\((\d+)\)\s*$", rest)
    if split and tail and int(split.group(1)) + int(split.group(2)) == int(tail.group(1)):
        value = float(tail.group(1))
    return re.sub(r"\s", "", m.group(1)), value


def isbn13(asin):
    """検索結果の項目ID（ISBN-10）をISBN-13にする。ISBN-10でなければそのまま返す。"""
    if not re.fullmatch(r"\d{9}[\dX]", asin):
        return asin
    body = "978" + asin[:9]
    check = (10 - sum(int(d) * (3 if i % 2 else 1) for i, d in enumerate(body)) % 10) % 10
    return body + str(check)


def parse_items(html):
    """
    検索結果ページのHTMLから本の一覧を抽出する。絞り込みはここでは行わない。

    ページ構造（実データで確認済み）：
      <li class="item" id="ISBN-10">
        <div class="Types"><span class="type type-tag">コミック</span></div>
        <div class="name">ONE PIECE 115</div>
        <div class="sab"><span>26年7月3日(金)</span><span>尾田栄一郎</span></div>
        <div class="sab"><span>ジャンプコミックス</span><span>集英社</span></div>
      </li>
    """
    out = []
    for li in BeautifulSoup(html, "html.parser").select("li.item"):
        name, sabs = li.select_one(".name"), li.select(".sab")
        if not name or not sabs:
            continue
        spans = [s.get_text(strip=True) for d in sabs for s in d.find_all("span")]
        dm = DATE_PATTERN.match(spans[0]) if spans else None
        if not dm:
            continue
        label_span = sabs[1].find("span") if len(sabs) > 1 else None
        tag = li.select_one(".type-tag")
        out.append({
            "name": name.get_text(strip=True),
            "category": tag.get_text(strip=True) if tag else "",
            "date": f"{2000 + int(dm.group(1))}/{int(dm.group(2)):02d}/{int(dm.group(3)):02d}",
            "isbn": isbn13(li.get("id", "")),
            "label": label_span.get_text(strip=True) if label_span else "",
            "spans": spans[1:],  # 先頭は著者
        })
    return out


def fetch_search_items(genre, title, page=1):
    """検索結果の本の一覧を返す。取得できなければ RequestException を送出する。"""
    url = GENRE_URLS.get(genre, GENRE_URLS["コミック"]) + requests.utils.quote(simplify_title_for_search(title))
    if page > 1:
        url += f"&p={page}"

    last_error = None
    MIN_VALID_RESPONSE_LEN = 1000  # 正常な検索結果ページは通常数万文字ある

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            # (connect timeout, read timeout) を分離。GitHub Actionsの
            # ランナーIPが一時的に不安定なことがあるため、接続段階は
            # 短めに切って早くリトライに回す。
            resp = requests.get(url, timeout=(10, 30), headers=HEADERS)
            resp.encoding = "utf-8"
            # 極端に短い応答はサーバー側の一時的なエラーとみなし、接続エラーと同じ扱いにする
            if len(resp.text) < MIN_VALID_RESPONSE_LEN:
                raise requests.exceptions.RequestException(f"応答が異常に短い（{len(resp.text)}文字）")
            return parse_items(resp.text)
        except requests.exceptions.RequestException as e:
            last_error = e
            print(f"⚠️ 接続失敗（{attempt}/{MAX_RETRIES}回目, {url}）: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SEC * attempt)

    raise last_error


# ------------------------------------------------------------
# 新刊の判定
# ------------------------------------------------------------
def base_label(label):
    """レーベル名の末尾に巻数が付く表記（"ノヴァコミックス 4"）をそろえる。"""
    return re.sub(r"\s*\d+$", "", label)


def fmt_vol(v):
    return str(int(v)) if v == int(v) else str(v)


def candidates(items, title, genre, longer_titles):
    """作品名で始まる対象分類の本に、目印・巻数・除外理由を付けて返す。"""
    allowed = ALLOWED_CATEGORIES.get(genre, ALLOWED_CATEGORIES["コミック"])
    for it in items:
        if it["category"] not in allowed:
            continue
        rest = split_after_title(it["name"], title)
        # 別の行の、より長い作品名に当たる本はその行のもの
        if rest is None or any(split_after_title(it["name"], t) is not None for t in longer_titles):
            continue
        sig, value = signature_and_volume(rest)
        excluded = "特装版・関連本" if NG_PATTERN.search(rest) else "外伝" if SPIN_PATTERN.search(rest) else None
        # 作品名の直後が区切り（空白・記号）か。別作品（作品名＋ーズ 等）を1巻と誤認しないため
        n = unicodedata.normalize("NFKC", it["name"])
        at = n.rfind(rest) if rest else len(n)
        boundary = rest == "" or not rest[0].isalnum() or (at > 0 and n[at - 1].isspace())
        yield dict(it, sig=sig, value=value, excluded=excluded, rest=rest, boundary=boundary)


def book_record(c):
    """次回の基準としてシートに保存する本の情報。"""
    return {"name": c["name"], "isbn": c["isbn"], "label": c["label"], "author": c["spans"][0] if c["spans"] else ""}


def new_result(status):
    return {"status": status, "new": [], "check": [], "save": None, "record": None, "date_change": None}


def judge_unrecorded(found, today):
    """B列が空の行（未発売の作品、または巻数をまだ記録していない作品）。"""
    books = [c for c in found if not c["excluded"] and (c["value"] is not None or c["boundary"])]
    if not books:
        return new_result("未発売")
    numbered = [c for c in books if c["value"] is not None]
    if all(c["value"] <= 1 for c in numbered):  # 1巻（番号なしを含む）だけがある
        pool, number = books, lambda c: 1.0
    else:  # 既刊がある作品を登録した場合は、本編（一番多い目印）の最新巻を基準にする
        main = max({c["sig"] for c in numbered}, key=lambda s: (sum(c["sig"] == s for c in numbered), s == ""))
        pool, number = [c for c in numbered if c["sig"] == main], lambda c: c["value"]
    if len({base_label(c["label"]) for c in pool}) > 1:  # 別の作品が混ざる可能性は人が確認する
        result = new_result("要確認")
        result["check"] = [f'{c["name"]}（{c["date"]}）レーベルが複数 {c["label"]}' for c in pool]
        return result
    pick = (min(pool, key=lambda c: c["date"]) if all(number(c) == 1 for c in pool)
            else max(pool, key=number))
    entry = dict(book_record(pick), vol=fmt_vol(number(pick)), date=pick["date"])
    result = new_result("新刊あり")
    result["save"] = book_record(pick)
    if number(pick) == 1 or pick["date"] >= today:  # 過去の巻は記録するだけで通知しない
        result["new"] = [entry]
    else:
        result.update(status="記録のみ", record=entry)
    return result


def judge(title, vol, cdate, genre, longer_titles, saved, today, fetch):
    """
    1行分の判定。vol/cdate はシートのB・C列、saved は前回保存した本（なければ None）、
    fetch(page) は検索結果の本の一覧を返す関数。

    本編の新刊の条件：作品名で始まる／レーベルが同じ／特装版・関連本・外伝でない／
    巻数の直前の言葉（目印）が基準の本と同じ／基準の本より発売日が新しい。
    """
    items = fetch(1)
    found = list(candidates(items, title, genre, longer_titles))
    try:
        current = float(vol)
    except ValueError:
        current = None

    # 保存した本は、巻数がB列と一致するときだけ使う（手動でB列を直した場合は推測し直す）
    saved_rest = split_after_title(saved["name"], title) if saved else None
    if saved_rest is None or current is None or (signature_and_volume(saved_rest)[1] or 1.0) != current:
        saved = None

    def read_page2():
        if len(items) >= 20:
            found.extend(candidates(fetch(2), title, genre, longer_titles))

    if not vol:
        read_page2()
        return judge_unrecorded(found, today)

    def days_apart(c):
        try:
            return abs((datetime.strptime(c["date"], "%Y/%m/%d") - datetime.strptime(cdate, "%Y/%m/%d")).days)
        except ValueError:
            return 0

    def find_anchor():
        if saved:
            for c in found:
                if (saved["isbn"] and c["isbn"] == saved["isbn"]) or c["name"] == saved["name"]:
                    return dict(c, from_site=True)
            sig, value = signature_and_volume(saved_rest)  # 検索結果に無くても書名から基準を復元する
            return {"name": saved["name"], "isbn": saved["isbn"], "date": cdate, "sig": sig, "value": value,
                    "rest": saved_rest, "label": saved["label"], "spans": [saved["author"]], "from_site": False}
        clean = [c for c in found if not c["excluded"] and c["sig"] is not None]
        for cond in (
            lambda c: c["date"] == cdate and c["value"] == current,
            lambda c: c["date"] == cdate,
            lambda c: c["value"] == current,
        ):
            hit = [c for c in clean if cond(c)]
            if hit:
                return min(hit, key=days_apart)
        return None

    anchor = find_anchor()
    if not saved and (not anchor or anchor["date"] != cdate) and len(items) >= 20:
        read_page2()
        anchor = find_anchor()
    if not anchor:
        result = new_result("目印未特定")
        result["check"] = ["検索結果が0件" if not items else "基準にできる本がない（作品名・巻数・発売日を確認）"]
        return result

    alt_sig = None
    if anchor["sig"] is None:  # 番号のない1巻が基準の場合
        alt_sig = re.sub(r"\s", "", anchor.get("rest", ""))
        anchor = dict(anchor, sig="", value=1.0)

    result = new_result("新刊なし")
    base = cdate
    if saved and anchor["from_site"] and anchor["date"] != cdate:  # 保存した本の発売日が動いた
        base = anchor["date"]
        result["date_change"] = (cdate, base)

    author = anchor["spans"][0] if anchor["spans"] else ""
    newer = sorted((c for c in found if base and c["date"] > base and c["isbn"] != anchor["isbn"]),
                   key=lambda c: c["date"])
    seen, high = set(), current
    for c in newer:
        if c["excluded"]:
            continue
        line = f'{c["name"]}（{c["date"]}）'
        same_label = base_label(anchor["label"]) in map(base_label, [c["label"], *c["spans"]])
        # レーベルが変わっても、同じ著者・同じ目印の次の巻なら続きとみなす
        moved = (c["sig"] == anchor["sig"] and author in c["spans"] and c["value"] is not None
                 and high is not None and high < c["value"] <= high + 1)
        if not same_label and not moved:
            continue
        if c["value"] is None:
            result["check"].append(f"{line}巻数なし")
        elif c["sig"] not in (anchor["sig"], alt_sig):
            result["check"].append(f"{line}本編の目印が違う")
        elif current is not None and c["value"] <= current:
            result["check"].append(f"{line}巻数が戻る")
        elif high is not None and c["value"] > high + 5:
            result["check"].append(f"{line}巻数が飛ぶ")
        elif c["value"] not in seen:
            seen.add(c["value"])
            high = c["value"] if high is None else max(high, c["value"])
            result["new"].append(dict(book_record(c), vol=fmt_vol(c["value"]), date=c["date"]))

    if result["new"]:
        result["status"] = "新刊あり"
        top = max(result["new"], key=lambda n: float(n["vol"]))
        result["save"] = {k: top[k] for k in ("name", "isbn", "label", "author")}
    else:
        if result["check"]:
            result["status"] = "要確認"
        if anchor["value"] == current:  # 巻数が一致する本だけを基準として保存する
            result["save"] = book_record(anchor)
    return result


# ------------------------------------------------------------
# 通知・カレンダー
# ------------------------------------------------------------
def format_volume_label(volume):
    """
    B列の巻数表記を表示用ラベルに変換する。
    数字のみ（通常の巻数）なら "第N巻"、それ以外（Ep.14 / Alter.2 のような
    特殊なシリーズ表記）はそのままの文字列を使う。
    """
    s = str(volume).strip()
    return f"第{s}巻" if re.fullmatch(r"\d+(\.\d+)?", s) else s


# Discord埋め込みの色（個別の新刊検知通知と、本日発売ダイジェストを
# 見分けやすくするため、固定色で使い分ける。ジャンルごとの色分けは
# 送信パターンが増えるため採用しない）
EMBED_COLOR_NEW_RELEASE = 0xE67E22  # オレンジ：新刊を検知した個別通知
EMBED_COLOR_DAILY_DIGEST = 0x3498DB  # 青：本日発売ダイジェスト
EMBED_COLOR_MANUAL_REVIEW_DIGEST = 0x9B59B6  # 紫：月次の手動確認まとめ


def send_discord(title, volume, release_date_str, genre, headline="新刊の発売情報を見つけました！"):
    if not DISCORD_WEBHOOK_URL:
        return

    emoji = "📚"
    if genre == "コミック":
        emoji = "🎨"
    elif genre in ("ライトノベル", "文庫"):
        emoji = "📖"

    volume_label = format_volume_label(volume)
    embed = {
        "title": f"{emoji} {headline}",
        "color": EMBED_COLOR_NEW_RELEASE,
        "fields": [
            {"name": "作品名", "value": title, "inline": False},
            {"name": "巻数", "value": volume_label, "inline": True},
            {"name": "発売日", "value": release_date_str, "inline": True},
            {"name": "ジャンル", "value": genre or "不明", "inline": True},
        ],
    }
    try:
        requests.post(DISCORD_WEBHOOK_URL, json={"embeds": [embed]}, timeout=10)
    except Exception as e:
        print(f"❌ Discord送信エラー: {e}")


def collect_todays_releases(calendar_service):
    """
    本日（JST）が開始日のカレンダーイベントを全件集める。
    スプレッドシートのC列（1タイトルにつき1つの日付しか持てない）ではなく
    カレンダー側を見ることで、2カ月連続発売のように複数巻の発売日が
    同時に検出された場合でも、各巻ごとの正しい発売日で拾えるようにしている。
    """
    today = now_jst().date()
    day_start = datetime(today.year, today.month, today.day, tzinfo=JST)
    day_end = day_start + timedelta(days=1)

    resp = (
        calendar_service.events()
        .list(
            calendarId=CALENDAR_ID,
            timeMin=day_start.astimezone(timezone.utc).isoformat(),
            timeMax=day_end.astimezone(timezone.utc).isoformat(),
            singleEvents=True,
        )
        .execute()
    )

    items = []
    for event in resp.get("items", []):
        props = event.get("extendedProperties", {}).get("private", {})
        title = props.get("title", "").strip()
        volume, genre = props.get("volume", ""), props.get("genre", "")
        if not title:
            # ver1.22より前に登録されたイベントはextendedPropertiesを持たないため、
            # 当時から入っていたdescription（作品タイトル：/巻数：/ジャンル：）から復元する
            fields = dict(
                line.split("：", 1) for line in event.get("description", "").splitlines() if "：" in line
            )
            title = fields.get("作品タイトル", "").strip()
            volume, genre = fields.get("巻数", "").strip(), fields.get("ジャンル", "").strip()
        if not title:
            continue  # このスクリプト以外が作成したイベント等は対象外
        items.append((title, volume, genre))
    return items


# S1セルに「本日分のダイジェストを送信済みか」の日付を記録する（再実行時の重複送信防止用）。
# GitHub Actionsのワークフローは接続不良時に新しいランナーでジョブ全体を
# 再実行するフォールバック構成になっており、ダイジェスト送信後に別の理由で
# ジョブが失敗すると、再実行時に同じ内容のダイジェストが再送されてしまう
# 問題があったため導入した。
DAILY_DIGEST_SENT_DATE_ROW = 1
DAILY_DIGEST_SENT_DATE_COL = 19  # S列（I1は検索列の見出しなので使わない）


def get_daily_digest_sent_date(rows):
    header_row = rows[0] if rows else []
    i = DAILY_DIGEST_SENT_DATE_COL - 1
    return header_row[i].strip() if len(header_row) > i else ""


def mark_daily_digest_sent(ws, date_str):
    batch_update_row(ws, DAILY_DIGEST_SENT_DATE_ROW, {DAILY_DIGEST_SENT_DATE_COL: force_text(date_str)})


def send_daily_digest(rows, ws, calendar_service):
    """
    本日発売予定の書籍を1回のDiscordメッセージにまとめて送信する。
    個別の新刊検知通知（send_discord）とは別に、本日が開始日のカレンダー
    イベントを全件集計する（スプレッドシートのG列フラグは問わない）。
    0件の日は送信しない。同じ日に二重実行（リトライ）されても再送しない
    よう、S1セルに送信済み日付を記録し、既に本日分を送信済みならスキップ
    する。
    """
    if not DISCORD_WEBHOOK_URL:
        return

    today_label = now_jst().strftime("%Y/%m/%d")
    if get_daily_digest_sent_date(rows) == today_label:
        print(f"⏭ 本日（{today_label}）分のダイジェストは送信済みのためスキップします。")
        return

    items = collect_todays_releases(calendar_service)
    if not items:
        print("📭 本日発売の書籍はありません（ダイジェスト送信をスキップ）。")
        mark_daily_digest_sent(ws, today_label)
        return

    lines = []
    for title, volume, genre in items:
        volume_label = format_volume_label(volume)
        genre_label = f"（{genre}）" if genre else ""
        lines.append(f"・{title} {volume_label}{genre_label}")

    # Discordの埋め込みdescription上限（4096文字）を考慮し、超える場合は
    # 複数の埋め込み（＝複数メッセージ）に分割して送信する。
    DISCORD_MAX_LEN = 3800
    chunks = []
    current, current_len = [], 0
    for line in lines:
        if current and current_len + len(line) + 1 > DISCORD_MAX_LEN:
            chunks.append(current)
            current, current_len = [], 0
        current.append(line)
        current_len += len(line) + 1
    if current:
        chunks.append(current)

    total_chunks = len(chunks)
    for idx, chunk_lines in enumerate(chunks, start=1):
        title_label = f"📚 本日（{today_label}）発売の書籍一覧"
        if total_chunks > 1:
            title_label += f"（{idx}/{total_chunks}）"
        embed = {
            "title": title_label,
            "description": "\n".join(chunk_lines),
            "color": EMBED_COLOR_DAILY_DIGEST,
            "footer": {"text": f"本日発売：{len(items)}件"},
        }
        try:
            requests.post(DISCORD_WEBHOOK_URL, json={"embeds": [embed]}, timeout=10)
        except Exception as e:
            print(f"❌ 本日発売ダイジェスト送信エラー: {e}")

    mark_daily_digest_sent(ws, today_label)
    print(f"📬 本日発売ダイジェストを送信しました（{len(items)}件、{len(chunks)}メッセージ）。")


def collect_manual_review_items(rows):
    """H列（手動更新フラグ）がTRUEの行を全件集める（自動検索で解決できず、手動確認が必要な行）"""
    items = []
    for row in rows[1:]:
        title = row[0].strip() if len(row) > 0 else ""
        genre = row[3].strip() if len(row) > 3 else ""
        is_manual_flag = row[7].strip() if len(row) > 7 else ""
        if not title or not is_true(is_manual_flag):
            continue
        items.append((title, genre))
    return items


def build_manual_search_url(title, genre):
    """
    スプレッドシートのHYPERLINK式（H列がTRUEの行に手動検索リンクを表示するための
    既存の運用フォーミュラ）と同じロジックで検索URLを組み立てる。
    ジャンルがGENRE_URLSの3種類（コミック/ライトノベル/文庫）以外の場合はNoneを返す。
    """
    base_url = GENRE_URLS.get(genre)
    if base_url is None:
        return None
    return base_url + requests.utils.quote(title)


def send_monthly_manual_review_digest(rows):
    """
    毎月1日の実行時にのみ、H列（手動更新フラグ）がTRUEのままの行を
    1回のDiscordメッセージにまとめて送信する（自動検索で解決できていない
    作品を月に1度まとめて手動確認してもらうため）。1日以外はスキップする。
    """
    if not DISCORD_WEBHOOK_URL:
        return

    today = now_jst()
    if today.day != 1:
        return

    items = collect_manual_review_items(rows)
    if not items:
        print("🗓 手動確認が必要な作品はありません（月次まとめ送信をスキップ）。")
        return

    lines = []
    for title, genre in items:
        url = build_manual_search_url(title, genre)
        genre_label = genre or "ジャンル不明"
        if url:
            lines.append(f"・[{title}]({url})（{genre_label}）")
        else:
            lines.append(f"・{title}（{genre_label}）")

    # Discordの埋め込みdescription上限（4096文字）を考慮し、超える場合は
    # 複数の埋め込み（＝複数メッセージ）に分割して送信する。
    DISCORD_MAX_LEN = 3800
    chunks = []
    current, current_len = [], 0
    for line in lines:
        if current and current_len + len(line) + 1 > DISCORD_MAX_LEN:
            chunks.append(current)
            current, current_len = [], 0
        current.append(line)
        current_len += len(line) + 1
    if current:
        chunks.append(current)

    month_label = today.strftime("%Y/%m")
    total_chunks = len(chunks)
    for idx, chunk_lines in enumerate(chunks, start=1):
        title_label = f"🔍 手動確認が必要な作品一覧（{month_label}）"
        if total_chunks > 1:
            title_label += f"（{idx}/{total_chunks}）"
        embed = {
            "title": title_label,
            "description": "\n".join(chunk_lines),
            "color": EMBED_COLOR_MANUAL_REVIEW_DIGEST,
            "footer": {"text": f"手動確認が必要な作品：{len(items)}件"},
        }
        try:
            requests.post(DISCORD_WEBHOOK_URL, json={"embeds": [embed]}, timeout=10)
        except Exception as e:
            print(f"❌ 月次手動確認まとめ送信エラー: {e}")

    print(f"📬 月次の手動確認まとめを送信しました（{len(items)}件、{len(chunks)}メッセージ）。")


def find_event(calendar_service, summary, release_date):
    """
    指定日にある、同じsummary（タイトル+巻数表記）のイベントを返す（なければNone）。
    重複登録の防止と、発売日が変わったときの予定の移動に使う。
    """
    day_start = datetime(release_date.year, release_date.month, release_date.day, tzinfo=JST)
    day_end = day_start + timedelta(days=1)

    resp = (
        calendar_service.events()
        .list(
            calendarId=CALENDAR_ID,
            timeMin=day_start.astimezone(timezone.utc).isoformat(),
            timeMax=day_end.astimezone(timezone.utc).isoformat(),
            singleEvents=True,
            q=summary,  # まず件数を絞るための全文検索（厳密一致は下でチェック）
        )
        .execute()
    )

    for event in resp.get("items", []):
        if event.get("summary", "") == summary:
            return event
    return None


def event_dates(release_date):
    return {
        "start": {"date": release_date.strftime("%Y-%m-%d")},
        "end": {"date": (release_date + timedelta(days=1)).strftime("%Y-%m-%d")},
    }


def move_calendar_event(calendar_service, title, volume, old_date, new_date, genre=None, isbn=""):
    """発売日が変わった本の予定を新しい日付へ移す。元の予定がなければ新しく作る。"""
    event = find_event(calendar_service, f"{title} {volume}", old_date)
    if event is None:
        create_calendar_event(calendar_service, title, volume, new_date, genre, isbn)
        return
    calendar_service.events().patch(
        calendarId=CALENDAR_ID, eventId=event["id"], body=event_dates(new_date)
    ).execute()


def create_calendar_event(calendar_service, title, volume, release_date, genre=None, isbn=""):
    volume_label = format_volume_label(volume)
    summary = f"{title} {volume}"

    if find_event(calendar_service, summary, release_date):
        print(f"⏭ カレンダーに同一イベントが既に存在するため登録をスキップ: {summary}")
        return False

    description_lines = [f"作品タイトル：{title}", f"巻数：{volume_label}"]
    if genre:
        description_lines.append(f"ジャンル：{genre}")
    event = {
        "summary": summary,
        "description": "\n".join(description_lines),
        **event_dates(release_date),
        # 本日発売ダイジェスト（send_daily_digest）がスプレッドシートのC列
        # （1タイトルにつき1つの日付しか持てない）ではなく、巻ごとに正しい
        # 日付を持つカレンダー側のイベントを直接読めるようにするための構造化データ。
        "extendedProperties": {
            "private": {"title": title, "volume": str(volume), "genre": genre or "", "isbn": isbn}
        },
    }
    calendar_service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
    return True


def cell(row, index):
    return str(row[index]).strip() if len(row) > index else ""


def is_true(value):
    return str(value).strip().upper() == "TRUE"


def parse_date_flexible(s):
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"日付形式を認識できません: {s}")


def force_text(value):
    """
    Sheets APIが「日付っぽい文字列」を勝手に日付型へ変換してしまい、
    セルの表示形式（たまたま日付のみ等）に従って情報が欠落する
    （例: "2026/08/09 12:27:09" → 読み込むと "2026/08/09" に時刻が消える）
    現象を防ぐため、先頭にアポストロフィを付けて強制的にプレーンテキスト
    として保存する。表示上アポストロフィそのものは見えない。
    """
    return f"'{value}"


def batch_update_row(ws, row_index, values):
    """
    1行分の複数セルを1回のAPI呼び出しでまとめて更新する。
    Sheets APIのクォータ（1分あたりのリクエスト数）を節約するため、
    update_cellを複数回呼ぶ代わりにこちらを使う。

    values: {列番号: 値} の辞書。例: {2: 5, 3: "2026/03/04", 5: "完了"}
    value_input_option="USER_ENTERED" にすることで、force_text() で
    付与したアポストロフィが「テキスト強制」として正しく解釈される。
    """
    cells = [gspread.Cell(row=row_index, col=col, value=value) for col, value in values.items()]
    ws.update_cells(cells, value_input_option="USER_ENTERED")


# ------------------------------------------------------------
# メイン処理
# ------------------------------------------------------------
def process_manual_rows(ws, calendar_service, rows):
    """J列（登録反映チェック）がONの行を最優先で処理する"""
    for i, row in enumerate(rows[1:], start=2):
        title = row[0].strip() if len(row) > 0 else ""
        is_manual_trigger = row[9].strip() if len(row) > 9 else ""
        if not title or not is_true(is_manual_trigger):
            continue

        try:
            current_latest_volume = row[1].strip() if len(row) > 1 else ""
            release_date_str = row[2].strip() if len(row) > 2 else ""
            genre = row[3].strip() if len(row) > 3 else ""
            is_notify_target = row[6].strip() if len(row) > 6 else ""

            if not current_latest_volume or not release_date_str:
                print(f"⚠️ Skip (Row {i}): B列またはC列が未入力のため手動スキップします。")
                continue

            release_date = parse_date_flexible(release_date_str)
            # B列は通常の巻数（数字）だけでなく、"Ep.14" "Alter.2" のような
            # 特殊なシリーズ表記もそのまま受け付ける（int変換は行わない）。
            matched_volume = current_latest_volume

            create_calendar_event(calendar_service, title, matched_volume, release_date, genre)

            if is_true(is_notify_target):
                send_discord(title, matched_volume, release_date.strftime("%Y/%m/%d"), genre)

            batch_update_row(
                ws,
                i,
                {5: "完了", 6: force_text(now_jst().strftime("%Y/%m/%d")), 10: False, COL_NOTE: ""},
            )

            print(f"✅ 【手動登録成功】 Row {i}: {title}（{matched_volume}）を処理しました。")

        except Exception as e:
            print(f"❌ 手動登録エラー（Row {i}）: {e}")


def select_rows_to_check(rows, daily_limit):
    """
    「最終更新日時（F列）が古い行」から優先的にdaily_limit件だけ選ぶ。
    新刊は数カ月おきにしか出ないため、毎日全件チェックする必要はない。
    1日あたりの処理件数に上限を設けることで、リクエスト数・実行時間・
    Sheets APIクォータを大きく節約しつつ、数日〜1週間程度で全件を
    一巡させる。

    戻り値には last_update も含める（選定理由をログで確認できるように）。
    """
    candidates = []
    parse_fail_examples = []
    for i, row in enumerate(rows[1:], start=2):
        title = row[0].strip() if len(row) > 0 else ""
        if not title:
            continue
        if is_true(row[7].strip() if len(row) > 7 else ""):
            continue  # H列TRUE（手動確認待ち）は自動検索が失敗する運命なのでスキップ
        last_update_str = row[5].strip() if len(row) > 5 else ""
        last_update = None
        for fmt in ("%Y/%m/%d", "%Y/%m/%d %H:%M:%S"):
            try:
                last_update = datetime.strptime(last_update_str, fmt)
                break
            except ValueError:
                continue
        if last_update is None:
            last_update = datetime.min  # 未チェックの行は最優先で処理する
            # 空文字ではないのにパースに失敗した場合は、実際に読み込んだ
            # 生の値をログに出す（フォーマットの想定違いを特定するため）
            if last_update_str and len(parse_fail_examples) < 5:
                parse_fail_examples.append((title, last_update_str))
        candidates.append((last_update, i, row))

    if parse_fail_examples:
        print(f"    ⚠️ F列のパースに失敗した行があります（未チェック扱いになります）。実際に読み込んだ値の例:")
        for title, raw_value in parse_fail_examples:
            print(f"        ・{title}: 読み込んだF列の値={raw_value!r}")

    candidates.sort(key=lambda x: x[0])  # 古い順（未チェック優先）
    return [(i, row, last_update) for last_update, i, row in candidates[:daily_limit]]


def touch_manual_review_rows(ws, rows):
    """
    H列（手動確認フラグ）がTRUEの行はサイト検索をスキップする代わりに、
    F列（最終更新日時）だけ今日の日付に更新する。検索はせずF列だけ
    触っておくことで、優先度付け（select_rows_to_check）で毎回
    「未チェックの最古行」扱いになり続けるのを防ぐ。
    """
    today_str = force_text(now_jst().strftime("%Y/%m/%d"))
    for i, row in enumerate(rows[1:], start=2):
        title = row[0].strip() if len(row) > 0 else ""
        if not title or not is_true(row[7].strip() if len(row) > 7 else ""):
            continue
        batch_update_row(ws, i, {6: today_str})


# シートの列番号（1始まり）
COL_VOLUME, COL_DATE, COL_STATUS, COL_CHECKED = 2, 3, 5, 6
SAVED_COLUMNS = {"name": 11, "isbn": 12, "label": 13, "author": 14}  # K〜N
COL_NOTE = 15  # O


def read_saved_book(row):
    """前回保存した本（K〜N列）。K列が空ならNone。"""
    book = {key: cell(row, col - 1) for key, col in SAVED_COLUMNS.items()}
    return book if book["name"] else None


def saved_cells(book):
    return {col: force_text(book[key]) if book[key] else "" for key, col in SAVED_COLUMNS.items()}


def to_cell_volume(volume):
    value = float(volume)
    return int(value) if value.is_integer() else value


def apply_result(ws, calendar_service, row_index, row, result):
    """判定結果を、カレンダー・Discord・シートに反映する。"""
    title, genre = cell(row, 0), cell(row, 3)
    notify = is_true(cell(row, 6))
    updates = {COL_STATUS: "完了", COL_CHECKED: force_text(now_jst().strftime("%Y/%m/%d"))}

    if result["date_change"]:
        old, new = result["date_change"]
        isbn = (read_saved_book(row) or {}).get("isbn", "")  # 日付が動いた本は、保存していた基準の本
        move_calendar_event(calendar_service, title, cell(row, 1),
                            parse_date_flexible(old), parse_date_flexible(new), genre, isbn)
        updates[COL_DATE] = force_text(new)
        if notify:
            send_discord(title, cell(row, 1), f"{old} → {new}", genre, headline="発売日が変わりました！")
        print(f"🗓 発売日変更: {title} {old} → {new}")

    for entry in result["new"]:
        release_date = parse_date_flexible(entry["date"])
        create_calendar_event(calendar_service, title, entry["vol"], release_date, genre, entry["isbn"])
        if notify:
            send_discord(title, entry["vol"], entry["date"], genre)
            print(f"🔔 登録＆Discord通知送信: {title} {entry['vol']}巻")
        else:
            print(f"📅 カレンダー登録のみ（通知OFF）: {title} {entry['vol']}巻")

    latest = max(result["new"], key=lambda n: float(n["vol"])) if result["new"] else result["record"]
    if latest:
        updates[COL_VOLUME] = to_cell_volume(latest["vol"])
        updates[COL_DATE] = force_text(latest["date"])
    if result["save"]:
        updates.update(saved_cells(result["save"]))
    updates[COL_NOTE] = " / ".join(result["check"][:3]) or ("未発売" if result["status"] == "未発売" else "")
    batch_update_row(ws, row_index, updates)

    if not result["new"] and not result["date_change"]:
        print(f"⏭ {result['status']}: {title}" + (f"（{updates[COL_NOTE]}）" if updates[COL_NOTE] else ""))


def process_auto_rows(ws, calendar_service, rows, daily_limit):
    touch_manual_review_rows(ws, rows)
    targets = select_rows_to_check(rows, daily_limit)
    print(f"📋 今回チェック対象: {len(targets)}件（全{len(rows) - 1}件中）")

    # 選定理由（前回チェック日）を出力し、なぜこの40件が選ばれたか
    # 後から確認できるようにする
    for i, row, last_update in targets:
        title = row[0].strip() if len(row) > 0 else ""
        last_update_display = "未チェック" if last_update == datetime.min else last_update.strftime("%Y/%m/%d")
        print(f"    ・{title}（前回チェック: {last_update_display}）")

    consecutive_network_failures = 0

    all_titles = {(cell(r, 3), cell(r, 0)) for r in rows[1:] if cell(r, 0)}

    for i, row, _ in targets:
        title, genre = cell(row, 0), cell(row, 3)
        # 同じジャンルで、この作品名より長い作品名の行がある場合、その本はそちらの行のもの
        longer_titles = [t for g, t in all_titles if g == genre and t != title
                         and normalize_title(t).startswith(normalize_title(title))]

        try:
            result = judge(title, cell(row, 1), cell(row, 2), genre, longer_titles, read_saved_book(row),
                           now_jst().strftime("%Y/%m/%d"),
                           lambda page: fetch_search_items(genre, title, page))
            consecutive_network_failures = 0  # 成功したのでリセット
            apply_result(ws, calendar_service, i, row, result)

        except requests.exceptions.RequestException as e:
            consecutive_network_failures += 1
            print(
                f"❌ 接続エラー（Row {i}, 連続{consecutive_network_failures}回目）: {e}"
            )
            if consecutive_network_failures >= MAX_CONSECUTIVE_NETWORK_FAILURES:
                print(
                    f"🛑 接続エラーが{MAX_CONSECUTIVE_NETWORK_FAILURES}回連続で発生したため処理を中断します"
                    f"（ランナーのネットワークが不調の可能性が高いです。ジョブを失敗させて"
                    f"ワークフロー側の再試行に委ねます）。"
                )
                raise TooManyNetworkFailures(
                    f"{MAX_CONSECUTIVE_NETWORK_FAILURES}回連続で接続エラーが発生しました"
                ) from e
        except Exception as e:
            print(f"❌ エラー（Row {i}）: {e}")

        time.sleep(REQUEST_INTERVAL_SEC)


def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    calendar_service = build("calendar", "v3", credentials=creds)

    print("⚡ 手動更新チェック（J列）の確認を開始します...")
    rows = ws.get_all_values()
    process_manual_rows(ws, calendar_service, rows)

    print(f"🔁 通常巡回処理を開始します（1日あたり最大{DAILY_LIMIT}件）...")
    rows = ws.get_all_values()  # 手動処理での更新を反映するため再取得
    process_auto_rows(ws, calendar_service, rows, DAILY_LIMIT)

    print("📬 本日発売の書籍ダイジェストを確認します...")
    rows = ws.get_all_values()  # 通常巡回での更新を反映するため再取得
    send_daily_digest(rows, ws, calendar_service)

    print("🗓 月次の手動確認まとめ（毎月1日のみ送信）を確認します...")
    send_monthly_manual_review_digest(rows)  # rowsはsend_daily_digestと同じもので良い（この間に書き込みはない）

    print("✅ 全処理が完了しました。")


if __name__ == "__main__":
    import sys

    try:
        main()
    except TooManyNetworkFailures as e:
        print(f"❌ 異常終了: {e}")
        sys.exit(1)  # ワークフロー側のフォールバックジョブに処理を委ねる