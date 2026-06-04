import os
import re
import time
import random
import unicodedata
import pandas as pd
from datetime import datetime, timedelta, timezone

# ==========================
# CONFIG
# ==========================

OUTPUT_FILE = "facebook_comments.xlsx"
MIN_WORDS   = 4
HEADLESS    = False

COLUMNS = ["ID", "platform", "created_date", "split",
           "text", "parent_text", "lang_type",
           "severity", "hate_type", "target", "implicit", "sarcasm"]

# ==========================
# Language Detection
# ==========================

BENGALI_CHAR_RE = re.compile(r'[\u0980-\u09FF]')

BANGLISH_KEYWORDS = {
    "ami","tumi","apni","se","tara","amra","tomra","apnara","amar","tomar",
    "apnar","tar","tader","amader","tomader","apnader","amake","tomake",
    "apnake","take","nij","nije","nijer","nijeke","nijeder","ache","achis",
    "achho","achhen","nei","nai","hobe","holo","hoye","hobi","hoen","hocche",
    "hoeche","thako","thakis","thaken","thaka","thakbe","thakbo","jabo","jabe",
    "jabi","jaoa","gelo","gechi","jacche","ashte","aste","asho","ashen","aso",
    "ashchi","ashbe","ashbo","kora","kori","koro","koren","korte","korchi",
    "korbe","korbo","korecho","korechi","koreni","korle","korleo","dekha",
    "dekhi","dekho","dekhen","dekhe","dekhchi","dekhbe","dekhbo","dekhecho",
    "dekheni","dekhle","bolo","boli","bolen","bolte","bolchi","bolbe","bolbo",
    "bolecho","boleni","bolechi","jani","jana","janish","janen","jante","janchi",
    "janbe","janbo","janecho","janeni","pari","parbo","paro","paren","parte",
    "parchi","parbe","parecho","pareni","lagbe","lage","lagche","laglo","lageni",
    "lagle","debo","dey","dao","den","deoa","diye","diyechi","diyeche","nebo",
    "nao","nen","niye","niyechi","niyeche","khao","khaish","khaen","khaoa",
    "khechi","kheye","khabe","khabo","shunte","shono","shonen","shunchi","shunbe",
    "shunbo","shunecho","bosho","bosen","bose","boschi","uthte","utho","uthen",
    "uthe","uthchi","pore","porchi","porbe","porbo","porechi","poreni","likho",
    "likhen","likhe","likhechi","likhbe","khelchi","khelo","khelen","khele",
    "shikho","shikhchi","shikhe","shikhbe","maro","mare","marchi","marbe","marbo",
    "haso","hase","haschi","hasbe","kando","kande","kanchi","kandbe","ghurchi",
    "ghuro","ghure","patha","pathai","pathao","pathaen","rakho","rakhe","rakhchi",
    "rakhbe","chao","chai","chaen","chawa","chechi","bujhi","bujho","bujhen",
    "bujhte","bujhchi","bujhbe","bujhechi","mone","moni","mono","monen","jan",
    "janen","ki","ke","keno","kothay","kothai","kotheke","kokhon","kivabe",
    "kibhabe","kon","konta","konti","kototuku","koto","kotota","kotogulo","kauke",
    "karo","kar","keu","kader","na","nah","haa","hya","han","ji","noi","noy",
    "ekdom","bilkul","thik","thikache","thikachhe","ar","kintu","tobe","tahole",
    "tai","karon","jodi","jokhon","jokhoni","ebong","ba","othoba","nahole","noyto",
    "nahoy","taopor","erpor","tarpor","kinthu","ebar","abar","tobuo","onek","khub",
    "aro","beshi","kom","ektu","ektukhani","khubi","shobcheye","sobcheye","shara",
    "sara","sotti","satyi","tokhon","ekhon","shei","sei","oi","ota","eta","emon",
    "eshob","shob","sob","shobai","shobkichu","sobkichu","shudhu","sudhu","tao",
    "porei","agei","ei","eka","ekjon","ekta","ekti","kono","kichhu","kichu",
    "kicchu","notun","purano","noya","naya","manush","lok","chele","meye","bhai",
    "apa","apu","dada","didi","mama","chacha","khalu","fufu","nana","nani","dadu",
    "dadi","ma","baba","maa","abbu","ammu","bondhu","bandhubi","bandu","kotha",
    "gaan","video","vedio","pic","picture","photo","channel","comment","like",
    "share","subscribe","sub","baari","bari","ghor","school","college","varsity",
    "haat","bazar","dokan","khana","khabar","pani","bhaat","roti","taka","poisha",
    "kaam","kaaj","cheez","jinish","somoy","shomoy","din","rat","sondha","bikel",
    "shokal","dupur","raat","valo","bhalo","kharap","sundor","shundor","mishti",
    "noshto","baje","oshadharon","oshadharan","boro","choto","lomba","khato","mota",
    "shukna","patla","shoja","kothin","shohoj","prothom","prothome","sheshe",
    "majhe","moddhe","thanda","gorom","naram","pobitro","bhalobasi","bhalobasha",
    "valobasa","valobasi","kosto","koshto","kasht","dukkho","dukkha","anondo",
    "ananda","khushi","miss","rag","raag","lojja","sharam","osthir","ostir",
    "pagol","pagla","paka","fatafati","joss","josh","darun","moja","maja",
    "nomoshkar","namaskar","salam","assalamu","alaikum","allah","eid","shubho",
    "subho","shubhechha","subhechha","dhonnobad","dhanybad","danyabad","shukriya",
    "sorry","maaf","vai","vaia","bro","sis","yaar","yar","re","da","lol","haha",
    "hihi","hehe","mashallah","masha","alhamdulillah","subhanallah","inshallah",
    "amin","ameen","waiting","plz","pls","please","help","request","darkar",
    "dorkar","link","daw","dhaka","chittagong","ctg","sylhet","rajshahi","khulna",
    "barishal","rangpur","mymensingh","bangladesh","bd","desh","bidesh","probash",
    "ekhane","okhane","shekhane","kothao","aj","aaj","kal","paroshhu","age",
    "shomoye","shoptah","mash","bochor","ghonta","minit","ek","dui","tin","char",
    "panch","choy","shat","ath","noy","dosh","sholo","bish","tish","de","ne",
    "ja","asha","bola","jawa",
}

EMOJI_RE = re.compile(
    "["
    "\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF"
    "\U0001F700-\U0001F77F\U0001F780-\U0001F7FF\U0001F800-\U0001F8FF"
    "\U0001F900-\U0001F9FF\U0001FA00-\U0001FA6F\U0001FA70-\U0001FAFF"
    "\U00002702-\U000027B0\U000024C2-\U0001F251\U0000200D\U0000FE0F"
    "]+", flags=re.UNICODE)

NOISE_RE = re.compile(
    r"^(Like|Reply|Comment|Share|Remove|Edit|Hide|Report|Follow|Unfollow|"
    r"See\s+translation|Translate|Top\s+fan|Author|Write\s+a\s+(comment|reply)|"
    r"\d+\s*(h|m|d|w|y|hr|min|sec|hour|minute|day|week|month|year)s?\.?|"
    r"\d+\s*(Like|Comment|Share|Reply)s?|[·•]+\s*\d*)$",
    re.IGNORECASE)


def remove_emojis(text):
    return EMOJI_RE.sub("", text)


def clean_text(raw):
    text = unicodedata.normalize("NFC", str(raw))
    text = remove_emojis(text)
    out = []
    for ch in text:
        cat = unicodedata.category(ch)
        if cat.startswith(("L","N","P","Z","M")) or ch in (" ","\n","\t","-","'",'"'):
            out.append(ch)
    return re.sub(r"\s+", " ", "".join(out)).strip()


def clean_comment(raw):
    """Strip FB UI noise lines from raw innerText."""
    if not raw:
        return ""
    good = []
    for line in raw.splitlines():
        line = line.strip()
        if line and not NOISE_RE.match(line):
            good.append(line)
    return clean_text(" ".join(good))


def has_min_words(text):
    return len(text.split()) >= MIN_WORDS


def today():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def parse_relative_date(raw):
    text = str(raw or "").strip().lower()
    if not text:
        return None

    now = datetime.now(timezone.utc)
    compact = re.sub(r"\s+", "", text)
    match = re.fullmatch(
        r"(\d+)(s|sec|secs|second|seconds|m|min|mins|minute|minutes|h|hr|hrs|hour|hours|d|day|days|w|wk|wks|week|weeks|mo|mon|month|months|y|yr|yrs|year|years)",
        compact,
    )
    if match:
        value = int(match.group(1))
        unit = match.group(2)
        if unit.startswith(("s", "sec")):
            dt = now - timedelta(seconds=value)
        elif unit.startswith(("m", "min")):
            dt = now - timedelta(minutes=value)
        elif unit.startswith(("h", "hr", "hour")):
            dt = now - timedelta(hours=value)
        elif unit.startswith(("d", "day")):
            dt = now - timedelta(days=value)
        elif unit.startswith(("w", "wk", "week")):
            dt = now - timedelta(weeks=value)
        elif unit.startswith(("mo", "mon", "month")):
            dt = now - timedelta(days=value * 30)
        else:
            dt = now - timedelta(days=value * 365)
        return dt.strftime("%Y-%m-%d")

    if text == "yesterday":
        return (now - timedelta(days=1)).strftime("%Y-%m-%d")
    if text in {"today", "just now"}:
        return now.strftime("%Y-%m-%d")
    return None


def format_date(raw):
    if not raw:
        return today()
    raw = str(raw).strip()
    if raw.isdigit():
        try:
            return datetime.fromtimestamp(int(raw), tz=timezone.utc).strftime("%Y-%m-%d")
        except Exception:
            pass
    relative = parse_relative_date(raw)
    if relative:
        return relative
    raw = re.sub(r"\bat\b", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"\s+", " ", raw).strip(" ,")
    for fmt in ("%Y-%m-%dT%H:%M:%S+0000", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ",
                "%B %d, %Y", "%d %B %Y", "%Y-%m-%d",
                "%A, %B %d, %Y %I:%M %p", "%A, %d %B %Y %I:%M %p",
                "%B %d, %Y %I:%M %p", "%d %B %Y %I:%M %p"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    current_year = datetime.now(timezone.utc).year
    for fmt in ("%B %d %I:%M %p", "%d %B %I:%M %p", "%B %d", "%d %B"):
        try:
            parsed = datetime.strptime(raw, fmt)
            return parsed.replace(year=current_year).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return today()


def detect_language(text):
    bn = BENGALI_CHAR_RE.findall(text)
    total = len([c for c in text if c.strip()])
    if total == 0:
        return None
    if len(bn) / total > 0.10:
        return "bn"
    words = re.findall(r"\b\w+\b", text.lower())
    if not words:
        return None
    if sum(1 for w in words if w in BANGLISH_KEYWORDS) / len(words) > 0.08:
        return "romanized"
    try:
        from langdetect import detect, DetectorFactory
        DetectorFactory.seed = 0
        d = detect(text)
        if d in ("en", "bn"):
            return d
    except Exception:
        pass
    return None


# ==========================
# Helpers
# ==========================

def assign_split():
    r = random.random()
    if r < 0.8:   return "train"
    elif r < 0.9: return "dev"
    return "test_new"


def load_existing():
    if os.path.exists(OUTPUT_FILE):
        df = pd.read_excel(OUTPUT_FILE)
        seen = set(zip(df["text"].astype(str), df["created_date"].astype(str)))
        return df, seen
    return None, set()


def get_next_id(df):
    if df is None or df.empty:
        return 1
    return int(str(df["ID"].iloc[-1]).split("_")[1]) + 1


def make_row(nid, text, parent_text, date, lang):
    return {
        "ID": f"BICH_{str(nid).zfill(4)}", "platform": "fb",
        "created_date": date, "split": assign_split(),
        "text": text, "parent_text": parent_text, "lang_type": lang,
        "severity": "", "hate_type": "", "target": "", "implicit": "", "sarcasm": "",
    }


def save_to_excel(new_rows):
    if not new_rows:
        print("⚠  No new rows to save."); return
    def nfc(v): return unicodedata.normalize("NFC", v) if isinstance(v, str) else v
    new_df = pd.DataFrame(new_rows, columns=COLUMNS)
    new_df["text"]        = new_df["text"].apply(nfc)
    new_df["parent_text"] = new_df["parent_text"].apply(nfc)
    if os.path.exists(OUTPUT_FILE):
        old_df = pd.read_excel(OUTPUT_FILE)
        for col in COLUMNS:
            if col not in old_df.columns: old_df[col] = ""
        combined = pd.concat([old_df[COLUMNS], new_df], ignore_index=True)
        combined.drop_duplicates(subset=["text","created_date"], inplace=True)
    else:
        combined = new_df
    combined[COLUMNS].to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Saved {len(new_rows)} new rows → {OUTPUT_FILE}")


# ==========================
# Selenium
# ==========================

def build_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    opts = Options()
    if HEADLESS: opts.add_argument("--headless=new")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--lang=en-US,en")
    opts.add_argument("--window-size=1366,900")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    except Exception:
        driver = webdriver.Chrome(options=opts)
    driver.execute_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
    return driver


def wait_for_login(driver):
    from selenium.webdriver.common.by import By
    print("\n🌐 Browser opened → log in to Facebook in the browser.")
    print("   Script continues automatically after login.\n")
    driver.get("https://www.facebook.com/")
    time.sleep(4)
    while True:
        try:
            driver.find_element(By.ID, "email"); time.sleep(2)
        except Exception:
            break
    print("✅ Login detected.\n")
    time.sleep(3)


def js_click(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def dismiss_dialogs(driver):
    from selenium.webdriver.common.by import By
    for txt in ["Accept all","Allow all cookies","Only allow essential cookies",
                "Close","Not now","Decline optional cookies"]:
        for btn in driver.find_elements(By.XPATH,
                f'//*[@role="button"][contains(.,"{txt}")]'):
            try: js_click(driver, btn); time.sleep(0.8)
            except: pass


def click_buttons_by_phrases(driver, phrases):
    return driver.execute_script("""
        const phrases = arguments[0];
        const buttons = Array.from(document.querySelectorAll(
            '[role="button"], button, a, div[tabindex="0"], div[tabindex="-1"], span[tabindex="0"]'
        ));
        let clicked = 0;
        for (const btn of buttons) {
            const text = (btn.innerText || btn.textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
            if (!text || btn.offsetParent === null) continue;
            if (!phrases.some(p => text.includes(p))) continue;
            if (btn.getAttribute('aria-disabled') === 'true') continue;
            btn.scrollIntoView({block: 'center'});
            btn.click();
            clicked += 1;
        }
        return clicked;
    """, [p.lower() for p in phrases])


def count_buttons_by_phrases(driver, phrases):
    return driver.execute_script("""
        const phrases = arguments[0];
        return Array.from(document.querySelectorAll(
            '[role="button"], button, a, div[tabindex="0"], div[tabindex="-1"], span[tabindex="0"]'
        ))
            .filter(btn => {
                const text = (btn.innerText || btn.textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                return text
                    && btn.offsetParent !== null
                    && btn.getAttribute('aria-disabled') !== 'true'
                    && phrases.some(p => text.includes(p));
            }).length;
    """, [p.lower() for p in phrases])


def page_height(driver):
    return driver.execute_script(
        "return Math.max(document.body.scrollHeight,"
        "document.documentElement.scrollHeight);")


def count_comment_articles(driver):
    return driver.execute_script(
        "return document.querySelectorAll('[role=\"article\"]').length;")


def comment_resource_count(driver):
    return driver.execute_script("""
        return performance.getEntriesByType('resource')
            .map(function(entry) { return entry.name || ''; })
            .filter(function(url) {
                return /comment|reply|ufi|feedback|plugins\\/feedback/i.test(url);
            }).length;
    """)


def scroll_comment_region(driver):
    return driver.execute_script("""
        function isVisible(el) {
            if (!el || !el.getBoundingClientRect) return false;
            var style = window.getComputedStyle(el);
            if (style.display === 'none' || style.visibility === 'hidden') return false;
            var rect = el.getBoundingClientRect();
            return rect.width > 20 && rect.height > 20;
        }

        function score(el) {
            var meta = [
                el.getAttribute('aria-label') || '',
                el.getAttribute('data-pagelet') || '',
                el.id || '',
                el.className || '',
                el.getAttribute('role') || ''
            ].join(' ').toLowerCase();
            var rect = el.getBoundingClientRect();
            var s = 0;
            if (/comment|reply|feed|ufi|dialog|focusable|story/.test(meta)) s += 12;
            s += Math.min(8, el.querySelectorAll('[role="article"]').length * 2);
            s += Math.min(6, el.querySelectorAll('[role="button"], button, a').length);
            if (rect.right > window.innerWidth * 0.45) s += 3;
            if (el.scrollHeight - el.clientHeight > 300) s += 4;
            return s;
        }

        var targets = Array.from(document.querySelectorAll('*'))
            .filter(function(el) {
                if (!isVisible(el)) return false;
                var style = window.getComputedStyle(el);
                var overflowY = style.overflowY || '';
                if (!/(auto|scroll|overlay)/.test(overflowY)) return false;
                return el.scrollHeight > el.clientHeight + 80;
            })
            .map(function(el) { return {el: el, score: score(el)}; })
            .sort(function(a, b) { return b.score - a.score; })
            .slice(0, 8);

        targets.forEach(function(entry) {
            var target = entry.el;
            try { target.focus && target.focus(); } catch (e) {}
            target.scrollTop = target.scrollHeight;
            target.dispatchEvent(new Event('scroll', {bubbles: true}));
            target.dispatchEvent(new WheelEvent('wheel', {deltaY: 1200, bubbles: true}));
        });

        if (document.scrollingElement) {
            document.scrollingElement.scrollTop = document.scrollingElement.scrollHeight;
        }
        window.scrollTo(0, Math.max(document.body.scrollHeight, document.documentElement.scrollHeight));
        return targets.length;
    """)


def scroll_to_bottom(driver):
    print("📜 Scrolling to bottom...")
    comment_phrases = [
        "view more comments",
        "see more comments",
        "more comments",
        "view previous comments",
        "previous comments",
    ]
    reply_phrases = [
        "1 reply",
        "view 1 reply",
        "see 1 reply",
        "view more replies",
        "see more replies",
        "more replies",
        "view previous replies",
        "previous replies",
        "replies",
    ]
    stale = 0
    no_growth = 0
    repeated_state = 0
    last_height = page_height(driver)
    last_articles = count_comment_articles(driver)
    last_resources = comment_resource_count(driver)
    last_signature = None
    rounds = 0
    while rounds < 45:
        rounds += 1
        clicked_comments = click_buttons_by_phrases(driver, comment_phrases)
        if clicked_comments:
            time.sleep(1.4)
        clicked_replies = click_buttons_by_phrases(driver, reply_phrases)
        if clicked_replies:
            time.sleep(1.1)

        scrolled_regions = scroll_comment_region(driver)
        driver.execute_script("window.scrollBy(0, 1200);")
        time.sleep(2.4)

        height = page_height(driver)
        articles = count_comment_articles(driver)
        resources = comment_resource_count(driver)
        remaining_comment_buttons = count_buttons_by_phrases(driver, comment_phrases)
        remaining_reply_buttons = count_buttons_by_phrases(driver, reply_phrases)
        print(
            f"  round={rounds} | height={height} | articles={articles} | "
            f"comment_requests={resources} | scrollers={scrolled_regions} | "
            f"comment_buttons={remaining_comment_buttons} | reply_buttons={remaining_reply_buttons}",
            end="\r",
        )
        grew = height > last_height or articles > last_articles or resources > last_resources
        if grew:
            no_growth = 0
        else:
            no_growth += 1

        signature = (
            height,
            articles,
            resources,
            scrolled_regions,
            clicked_comments,
            clicked_replies,
            remaining_comment_buttons,
            remaining_reply_buttons,
        )
        if signature == last_signature:
            repeated_state += 1
        else:
            repeated_state = 0
            last_signature = signature

        if (
            height <= last_height
            and articles <= last_articles
            and resources <= last_resources
            and clicked_comments == 0
            and clicked_replies == 0
            and remaining_comment_buttons == 0
            and remaining_reply_buttons == 0
        ):
            stale += 1
        else:
            stale = 0
            last_height = height
            last_articles = articles
            last_resources = resources

        if stale >= 8 or no_growth >= 8 or repeated_state >= 6:
            break

    print(f"\n  ↳ Done. Final height: {page_height(driver)}px")


def expand_replies(driver):
    print("💬 Expanding replies...")
    scroll_to_bottom(driver)
    print("  ↳ Done.")


# ==========================
# JS extraction — revised
# ==========================

EXTRACT_JS = """
return (function() {

    // ── 1. Helpers ────────────────────────────────────────────────────────────

    function normalizeText(value) {
        return (value || '').replace(/\\s+/g, ' ').trim();
    }

    function articleDepth(el) {
        var d = 0, p = el.parentElement;
        while (p) {
            if (p.getAttribute && p.getAttribute('role') === 'article') d++;
            p = p.parentElement;
        }
        return d;
    }

    // Get left offset of element relative to page — used to detect indentation
    function leftOffset(el) {
        var x = 0;
        while (el) { x += el.offsetLeft || 0; el = el.offsetParent; }
        return x;
    }

    function hasReplyAction(article) {
        var controls = article.querySelectorAll(
            '[role="button"], button, a, div[tabindex="0"], div[tabindex="-1"], span[tabindex="0"]'
        );
        for (var i = 0; i < controls.length; i++) {
            var txt = normalizeText(
                controls[i].innerText ||
                controls[i].textContent ||
                controls[i].getAttribute('aria-label') ||
                ''
            ).toLowerCase();
            if (/\\brepl(y|ies)\\b/.test(txt)) return true;
        }
        return false;
    }

    // Extract comment body text from a single article element
    function ownText(article) {
        var clone = article.cloneNode(true);
        // Strip nested articles (shouldn't exist in flat layout, but safety net)
        clone.querySelectorAll('[role="article"]').forEach(function(n) {
            n.parentNode && n.parentNode.removeChild(n);
        });
        // Strip UI chrome: like/reply buttons, reaction counts, timestamps
        clone.querySelectorAll([
            '[role="button"]',
            'a[href*="reaction"]',
            'span[aria-label]',
        ].join(',')).forEach(function(n) {
            // Only remove if short UI text (not the comment itself)
            var t = (n.innerText || '').trim();
            if (t.length < 20) { n.parentNode && n.parentNode.removeChild(n); }
        });

        // Priority 1: data-ad-preview="message"
        var msg = clone.querySelector('[data-ad-preview="message"]');
        if (msg) {
            var t = (msg.innerText || msg.textContent || '').trim();
            if (t.length > 2) return t;
        }

        // Priority 2: longest [dir="auto"] block
        var autos = Array.from(clone.querySelectorAll('[dir="auto"]'));
        autos.sort(function(a, b) {
            return (b.innerText || b.textContent || '').trim().length -
                   (a.innerText || a.textContent || '').trim().length;
        });
        for (var i = 0; i < autos.length; i++) {
            var t = (autos[i].innerText || autos[i].textContent || '').trim();
            if (t.length > 2) return t;
        }

        // Priority 3: [lang] attribute elements
        var langEls = clone.querySelectorAll('[lang]');
        for (var j = 0; j < langEls.length; j++) {
            var t = (langEls[j].innerText || langEls[j].textContent || '').trim();
            if (t.length > 2) return t;
        }

        return (clone.innerText || clone.textContent || '').trim();
    }

    function getDate(article) {
        var abbrs = article.querySelectorAll('abbr[data-utime]');
        if (abbrs.length) return abbrs[0].getAttribute('data-utime') || '';
        var times = article.querySelectorAll('time[datetime]');
        if (times.length) return times[0].getAttribute('datetime') || '';
        var dateNodes = article.querySelectorAll('a[aria-label], a[title], span[title], span[aria-label]');
        for (var k = 0; k < dateNodes.length; k++) {
            var label = dateNodes[k].getAttribute('aria-label') || dateNodes[k].getAttribute('title') || '';
            if (/\d{4}|\b(today|yesterday)\b|\b\d+\s*(m|min|h|d|w|mo|y)\b/i.test(label))
                return label;
        }
        var links = article.querySelectorAll('a[aria-label]');
        for (var i = 0; i < links.length; i++) {
            var lbl = links[i].getAttribute('aria-label') || '';
            if (/\d{4}|\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i.test(lbl))
                return lbl;
        }
        var titled = article.querySelectorAll('[title]');
        for (var j = 0; j < titled.length; j++) {
            var ttl = titled[j].getAttribute('title') || '';
            if (/\d{4}/.test(ttl) && ttl.length < 60) return ttl;
        }
        return '';
    }

    // ── 2. Collect all articles and compute depths ────────────────────────────
    var allArticles = Array.from(document.querySelectorAll('[role="article"]'));
    if (!allArticles.length) return [];

    var depths = allArticles.map(articleDepth);
    var positive = depths.filter(function(d) { return d > 0; });
    var rootDepth = positive.length ? Math.min.apply(null, positive) : 0;

    // ── 3. Separate post body (depth 0), root comments, and replies ───────────
    // Facebook's modern layout: ALL comments (root + replies) sit at the SAME
    // DOM depth. We distinguish them by LEFT OFFSET — replies are indented more.

    var commentArticles = allArticles.filter(function(a, i) {
        return positive.length ? depths[i] >= rootDepth : hasReplyAction(a);
    });

    if (!commentArticles.length) {
        commentArticles = allArticles.slice(1);
    }
    if (!commentArticles.length) return [];

    // Compute left offsets for all comment articles
    var offsets = commentArticles.map(leftOffset);
    var minOffset = Math.min.apply(null, offsets);

    // Root comments sit at minOffset; replies are indented further
    // Use a tolerance of 20px to handle slight measurement variance
    var INDENT_THRESHOLD = 20;

    // ── 4. Group by sequential scan: each root starts a new thread ───────────
    var threads = [];
    var currentThread = null;

    commentArticles.forEach(function(a, i) {
        var offset = offsets[i];
        var isRoot = (offset - minOffset) < INDENT_THRESHOLD;

        if (isRoot) {
            // Start a new thread
            currentThread = {
                root_text: ownText(a),
                root_date: getDate(a),
                replies: []
            };
            threads.push(currentThread);
        } else {
            // This is a reply — attach to the most recent root thread
            if (!currentThread) {
                // Edge case: reply before any root (shouldn't happen, but safe)
                currentThread = { root_text: '', root_date: '', replies: [] };
                threads.push(currentThread);
            }
            var t = ownText(a);
            if (t) {
                currentThread.replies.push({ text: t, date: getDate(a) });
            }
        }
    });

    return threads;
})();
"""


def scrape_comments(driver):
    print("🔍 Extracting comments via JS...")
    try:
        result = driver.execute_script(EXTRACT_JS)
    except Exception as e:
        print(f"  ↳ JS error: {e}"); return []

    if not result:
        debug = driver.execute_script("""
            return (function(){
                var all = document.querySelectorAll('[role="article"]');
                var info = [];
                all.forEach(function(a){
                    var d=0, p=a.parentElement;
                    while(p){if(p.getAttribute&&p.getAttribute('role')==='article')d++;p=p.parentElement;}
                    var x=0, el=a;
                    while(el){x+=el.offsetLeft||0; el=el.offsetParent;}
                    var t=(a.innerText||'').substring(0,60).replace(/\\n/g,' ');
                    info.push({depth:d, offsetLeft:x, snippet:t});
                });
                return info;
            })();
        """)
        print(f"  ↳ DEBUG — articles on page ({len(debug)} found):")
        for item in debug[:15]:
            print(f"      depth={item['depth']}  offset={item['offsetLeft']}  {repr(item['snippet'])}")
        return []

    total_roots   = len(result)
    total_replies = sum(len(t.get("replies", [])) for t in result)
    print(f"  ↳ Root comments: {total_roots} | Total replies: {total_replies}")
    if result:
        s = result[0]
        print(f"  ↳ Sample root:    {repr(s['root_text'][:80])}")
        print(f"  ↳ Sample reply 0: {repr(s['replies'][0] if s['replies'] else 'none')}")
    return result


# ==========================
# Main
# ==========================

def fetch_comments(post_url):
    post_url = post_url.strip()
    post_url = post_url.replace("web.facebook.com", "www.facebook.com")

    old_df, seen = load_existing()
    nid          = get_next_id(old_df)
    new_rows     = []
    stats        = {"added": 0, "short": 0, "lang": 0, "dup": 0, "empty": 0}

    driver = build_driver()
    try:
        wait_for_login(driver)

        print(f"🌐 Opening: {post_url}")
        driver.get(post_url)
        time.sleep(6)
        dismiss_dialogs(driver)
        time.sleep(2)

        resolved = driver.current_url
        if resolved != post_url:
            print(f"  ↳ Resolved to: {resolved}")

        scroll_to_bottom(driver)

        driver.execute_script(
            "window.scrollTo(0,Math.max(document.body.scrollHeight,"
            "document.documentElement.scrollHeight));")
        time.sleep(3)

        threads = scrape_comments(driver)

    finally:
        driver.quit()

    print(f"\n📊 Processing {len(threads)} thread(s)...")

    for thread in threads:
        root_text = clean_comment(thread.get("root_text", ""))

        # ── Only replies become rows; root comment goes into parent_text ──────
        for reply in thread.get("replies", []):
            rep_text = clean_comment(reply.get("text", ""))
            rep_date = format_date(reply.get("date", ""))

            if not rep_text:
                stats["empty"] += 1; continue
            if not has_min_words(rep_text):
                stats["short"] += 1; continue
            if (rep_text, rep_date) in seen:
                stats["dup"] += 1; continue

            lang = detect_language(rep_text)
            if lang is None:
                stats["lang"] += 1; continue

            new_rows.append(make_row(nid, rep_text, root_text, rep_date, lang))
            seen.add((rep_text, rep_date))
            nid += 1
            stats["added"] += 1

    print(f"📊 Added:{stats['added']} | Lang:{stats['lang']} | "
          f"Dup:{stats['dup']} | Short:{stats['short']} | Empty:{stats['empty']}")
    save_to_excel(new_rows)


# ==========================
# CLI
# ==========================

if __name__ == "__main__":
    print("📘 Facebook Comment Scraper")
    print("─" * 45)
    print("pip install selenium webdriver-manager pandas openpyxl langdetect")
    print("─" * 45)
    post_link = input("📌 Facebook post URL: ").strip()
    fetch_comments(post_link)
