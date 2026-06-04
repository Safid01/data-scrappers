import os
import re
import random
import unicodedata
import pandas as pd
from datetime import datetime
from langdetect import detect, DetectorFactory

DetectorFactory.seed = 0

# ==========================
# CONFIG
# ==========================

OUTPUT_FILE = "youtube_comments.xlsx"
MIN_WORDS = 4

# ==========================
# Language Detection
# ==========================

BENGALI_CHAR_RE = re.compile(r'[\u0980-\u09FF]')

BANGLISH_KEYWORDS = {
    # ---------- Pronouns ----------
    "ami", "tumi", "apni", "se", "tara", "amra", "tomra", "apnara",
    "amar", "tomar", "apnar", "tar", "tader", "amader", "tomader",
    "apnader", "amake", "tomake", "apnake", "take",
    "nij", "nije", "nijer", "nijeke", "nijeder",

    # ---------- Common verbs (roots + conjugations) ----------
    "ache", "achis", "achho", "achhen", "nei", "nai",
    "hobe", "holo", "hoye", "hobi", "hoen", "hocche", "hoeche",
    "thako", "thakis", "thaken", "thaka", "thakbe", "thakbo",
    "jabo", "jabe", "jabi", "jaoa", "gelo", "gechi", "jacche",
    "ashte", "aste", "asho", "ashen", "aso", "ashchi", "ashbe", "ashbo",
    "kora", "kori", "koro", "koren", "korte", "korchi", "korbe", "korbo",
    "korecho", "korechi", "koreni", "korle", "korleo",
    "dekha", "dekhi", "dekho", "dekhen", "dekhe", "dekhchi", "dekhbe",
    "dekhbo", "dekhecho", "dekheni", "dekhle",
    "bolo", "boli", "bolen", "bolte", "bolchi", "bolbe", "bolbo",
    "bolecho", "boleni", "bolechi",
    "jani", "jana", "janish", "janen", "jante", "janchi", "janbe",
    "janbo", "janecho", "janeni",
    "pari", "parbo", "paro", "paren", "parte", "parchi", "parbe",
    "parecho", "pareni",
    "lagbe", "lage", "lagche", "laglo", "lageni", "lagle",
    "debo", "dey", "dao", "den", "deoa", "diye", "diyechi", "diyeche",
    "nebo", "nao", "nen", "niye", "niyechi", "niyeche",
    "khao", "khaish", "khaen", "khaoa", "khechi", "kheye", "khabe", "khabo",
    "shunte", "shono", "shonen", "shunchi", "shunbe", "shunbo", "shunecho",
    "bosho", "bosen", "bose", "boschi",
    "uthte", "utho", "uthen", "uthe", "uthchi",
    "paro", "pore", "porchi", "porbe", "porbo", "porechi", "poreni",
    "likho", "likhen", "likhe", "likhechi", "likhbe",
    "khelchi", "khelo", "khelen", "khele",
    "shikho", "shikhchi", "shikhe", "shikhbe",
    "maro", "mare", "marchi", "marbe", "marbo",
    "haso", "hase", "haschi", "hasbe",
    "kando", "kande", "kanchi", "kandbe",
    "ghurchi", "ghuro", "ghure",
    "patha", "pathai", "pathao", "pathaen",
    "rakho", "rakhe", "rakhchi", "rakhbe",
    "chao", "chai", "chaen", "chawa", "chechi",
    "bujhi", "bujho", "bujhen", "bujhte", "bujhchi", "bujhbe", "bujhechi",
    "mone", "moni", "mono", "monen",
    "jan", "janen",

    # ---------- Question words ----------
    "ki", "ke", "keno", "kothay", "kothai", "kotheke", "kokhon",
    "kivabe", "kibhabe", "kon", "konta", "konti", "kototuku", "koto",
    "kotota", "kotogulo", "kauke", "karo", "kar", "keu", "kader",

    # ---------- Negation / affirmation ----------
    "na", "nah", "haa", "hya", "han", "ji", "noi", "noy",
    "ekdom", "bilkul", "thik", "thikache", "thikachhe",

    # ---------- Conjunctions / connectors ----------
    "ar", "kintu", "tobe", "tahole", "tai", "karon", "jodi",
    "jokhon", "jokhoni", "ebong", "ba", "othoba",
    "nahole", "noyto", "nahoy", "taopor", "erpor",
    "tarpor", "kinthu", "ebar", "abar", "tobuo",

    # ---------- Adverbs / intensifiers ----------
    "onek", "khub", "aro", "beshi", "kom", "ektu", "ektukhani",
    "khubi", "shobcheye", "sobcheye", "shara", "sara",
    "sotti", "satyi", "tokhon", "ekhon", "shei",
    "sei", "oi", "ota", "eta", "emon", "eshob", "shob", "sob",
    "shobai", "shobkichu", "sobkichu", "shudhu", "sudhu",
    "tao", "porei", "agei",

    # ---------- Demonstratives / determiners ----------
    "ei", "eka", "ekjon", "ekta", "ekti", "kono",
    "kichhu", "kichu", "kicchu", "notun", "purano", "noya", "naya",

    # ---------- Common nouns ----------
    "manush", "lok", "chele", "meye", "bhai", "apa", "apu",
    "dada", "didi", "mama", "chacha", "khalu", "fufu", "nana",
    "nani", "dadu", "dadi", "ma", "baba", "maa", "abbu", "ammu",
    "bondhu", "bandhubi", "bandu",
    "kotha", "gaan", "video", "vedio", "pic", "picture", "photo",
    "channel", "comment", "like", "share", "subscribe", "sub",
    "baari", "bari", "ghor", "school", "college", "varsity",
    "haat", "bazar", "dokan", "khana", "khabar", "pani", "bhaat",
    "roti", "taka", "poisha", "kaam", "kaaj", "cheez", "jinish",
    "somoy", "shomoy", "din", "rat", "sondha", "bikel", "shokal",
    "dupur", "raat",

    # ---------- Adjectives ----------
    "valo", "bhalo", "kharap", "sundor", "shundor",
    "mishti", "noshto", "baje", "oshadharon", "oshadharan",
    "boro", "choto", "lomba", "khato", "mota", "shukna", "patla",
    "shoja", "kothin", "shohoj", "prothom", "prothome",
    "sheshe", "majhe", "moddhe",
    "thanda", "gorom", "naram", "pobitro",

    # ---------- Emotions / reactions ----------
    "bhalobasi", "bhalobasha", "valobasa", "valobasi",
    "kosto", "koshto", "kasht", "dukkho", "dukkha", "anondo",
    "ananda", "khushi", "miss",
    "rag", "raag", "lojja", "sharam",
    "osthir", "ostir", "pagol", "pagla", "paka",
    "fatafati", "joss", "josh",
    "darun", "moja", "maja",

    # ---------- Greetings / closings ----------
    "nomoshkar", "namaskar", "salam", "assalamu", "alaikum",
    "allah", "eid", "shubho", "subho",
    "shubhechha", "subhechha",
    "dhonnobad", "dhanybad", "danyabad", "shukriya",
    "sorry", "maaf",

    # ---------- Internet / YouTube slang ----------
    "vai", "vaia", "bro", "sis",
    "yaar", "yar", "re", "da",
    "lol", "haha", "hihi", "hehe",
    "mashallah", "masha", "alhamdulillah", "subhanallah",
    "inshallah", "amin", "ameen",
    "waiting", "plz", "pls", "please", "help",
    "request", "darkar", "dorkar",
    "link", "daw",

    # ---------- Location words ----------
    "dhaka", "chittagong", "ctg", "sylhet", "rajshahi",
    "khulna", "barishal", "rangpur", "mymensingh",
    "bangladesh", "bd", "desh", "bidesh", "probash",
    "ekhane", "okhane", "shekhane", "kothao",

    # ---------- Time words ----------
    "aj", "aaj", "kal", "paroshhu", "age",
    "shomoye", "shoptah", "mash", "bochor",
    "ghonta", "minit",

    # ---------- Numbers ----------
    "ek", "dui", "tin", "char", "panch", "choy", "shat",
    "ath", "noy", "dosh", "sholo", "bish", "tish",

    # ---------- Verb roots typed alone ----------
    "de", "ne", "ja", "asha", "bola", "jawa",
}

# Matches all emoji / pictographic / symbol Unicode ranges
EMOJI_RE = re.compile(
    "["
    "\U0001F600-\U0001F64F"
    "\U0001F300-\U0001F5FF"
    "\U0001F680-\U0001F6FF"
    "\U0001F700-\U0001F77F"
    "\U0001F780-\U0001F7FF"
    "\U0001F800-\U0001F8FF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA00-\U0001FA6F"
    "\U0001FA70-\U0001FAFF"
    "\U00002702-\U000027B0"
    "\U000024C2-\U0001F251"
    "\U0000200D"
    "\U0000FE0F"
    "]+",
    flags=re.UNICODE,
)


def remove_emojis(text):
    """Remove emoji characters but keep all surrounding text."""
    return EMOJI_RE.sub("", text)


def clean_text(text):
    """NFC-normalize, remove emojis, then strip non-printable characters."""
    text = unicodedata.normalize("NFC", text)
    text = remove_emojis(text)
    cleaned = []
    for char in text:
        cat = unicodedata.category(char)
        if cat.startswith(('L', 'N', 'P', 'Z', 'M')) or char in (' ', '\n', '\t', '-', "'", '"'):
            cleaned.append(char)
    return re.sub(r'\s+', ' ', ''.join(cleaned)).strip()


def has_min_words(text, min_words=MIN_WORDS):
    return len(text.split()) >= min_words


def format_date(timestamp):
    try:
        return datetime.utcfromtimestamp(int(timestamp)).strftime("%Y-%m-%d")
    except Exception:
        return str(timestamp)


def detect_language(text):
    bengali_chars = BENGALI_CHAR_RE.findall(text)
    total_chars = len([c for c in text if c.strip()])
    if total_chars == 0:
        return None

    if len(bengali_chars) / total_chars > 0.10:
        return "bn"

    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return None

    banglish_hits = sum(1 for w in words if w in BANGLISH_KEYWORDS)
    if banglish_hits / len(words) > 0.08:
        return "romanized"

    try:
        detected = detect(text)
        if detected == "en":
            return "en"
        if detected == "bn":
            return "bn"
    except Exception:
        pass

    return None


# ==========================
# Helpers
# ==========================

def assign_split():
    r = random.random()
    if r < 0.8:
        return "train"
    elif r < 0.9:
        return "dev"
    return "test_new"


def load_existing_data():
    if os.path.exists(OUTPUT_FILE):
        df = pd.read_excel(OUTPUT_FILE)
        existing_pairs = set(zip(df["text"], df["created_date"]))
        return df, existing_pairs
    return None, set()


def get_next_id(df):
    if df is None or df.empty:
        return 1
    last_id = df["ID"].iloc[-1]
    return int(last_id.split("_")[1]) + 1


# ==========================
# Recursive reply collector
# ==========================

def collect_all_replies(comment, children_map, parent_text):
    """
    Recursively yield (reply, parent_text) for all descendants.
    parent_text is always the cleaned text of the TOP-LEVEL comment.
    """
    cid = comment.get("id", "")
    for child in children_map.get(cid, []):
        yield child, parent_text
        yield from collect_all_replies(child, children_map, parent_text)


# ==========================
# Main Fetcher (yt-dlp only)
# ==========================

def fetch_comments(video_url):
    try:
        import yt_dlp
    except ImportError:
        print("⚠ yt-dlp not installed. Run: pip install yt-dlp")
        return

    old_df, existing_pairs = load_existing_data()
    next_id = get_next_id(old_df)

    print("🔍 Fetching comments via yt-dlp...")

    ydl_opts = {
        'getcomments': True,
        'skip_download': True,
        'quiet': True,
        'no_warnings': True,
        'extractor_args': {'youtube': {'comment_sort': ['top']}},
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(video_url, download=False)
        comments = info.get("comments", [])

    print(f"📥 Total comments fetched: {len(comments)}")

    root_comments = []
    children_map = {}

    for c in comments:
        pid = c.get("parent", "root")
        if pid == "root":
            root_comments.append(c)
        else:
            children_map.setdefault(pid, []).append(c)

    print(f"📊 Top-level comments: {len(root_comments)}")

    new_rows = []
    stats = {"lang": 0, "dup": 0, "short": 0, "no_text": 0, "added": 0}

    for top_comment in root_comments:
        top_text = clean_text(top_comment.get("text", ""))
        top_id = top_comment.get("id", "")

        if top_id not in children_map:
            continue

        parent_text = top_text

        for reply, pt in collect_all_replies(top_comment, children_map, parent_text):
            reply_text = clean_text(reply.get("text", ""))

            if not reply_text:
                stats["no_text"] += 1
                continue

            if not has_min_words(reply_text):
                stats["short"] += 1
                continue

            reply_date = format_date(reply.get("timestamp", ""))

            if (reply_text, reply_date) in existing_pairs:
                stats["dup"] += 1
                continue

            lang = detect_language(reply_text)
            if lang is None:
                stats["lang"] += 1
                continue

            new_rows.append({
                "ID": f"BICH_{str(next_id).zfill(4)}",
                "platform": "yt",
                "created_date": reply_date,
                "split": assign_split(),
                "text": reply_text,
                "parent_text": pt,
                "lang_type": lang,
                "severity": "",
                "hate_type": "",
                "target": "",
                "implicit": "",
                "sarcasm": "",
            })
            existing_pairs.add((reply_text, reply_date))
            next_id += 1
            stats["added"] += 1

    print(
        f"📊 Results — Added: {stats['added']} | "
        f"Bad lang: {stats['lang']} | "
        f"Duplicates: {stats['dup']} | "
        f"Too short (<{MIN_WORDS} words): {stats['short']} | "
        f"Empty after clean: {stats['no_text']}"
    )
    save_to_excel(new_rows)


# ==========================
# Save to Excel
# ==========================

def save_to_excel(new_rows):
    if not new_rows:
        print("⚠ No new rows to save.")
        return

    def nfc(val):
        return unicodedata.normalize("NFC", val) if isinstance(val, str) else val

    new_df = pd.DataFrame(new_rows)
    for col in ["text", "parent_text"]:
        if col in new_df.columns:
            new_df[col] = new_df[col].apply(nfc)

    if os.path.exists(OUTPUT_FILE):
        old_df = pd.read_excel(OUTPUT_FILE)
        combined = pd.concat([old_df, new_df], ignore_index=True)
        combined.drop_duplicates(subset=["text", "created_date"], inplace=True)
    else:
        combined = new_df

    combined.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Saved {len(new_rows)} new rows to {OUTPUT_FILE}")


# ==========================
# CLI
# ==========================

if __name__ == "__main__":
    video_link = input("🎥 Enter YouTube video link: ").strip()
    fetch_comments(video_link)