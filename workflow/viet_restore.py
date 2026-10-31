# -*- coding: utf-8 -*-
"""
Xử lý chữ tiếng Việt bị hỏng trong báo cáo .xls
===============================================
Kiểm tra thực tế file "19 - 25.10.xls": KHÔNG phải TCVN3 (ABC). File được lưu theo bảng mã Windows-1252, nên mọi chữ
Việt nằm ngoài bảng mã này (ă, đ, ơ, ư, ạ, ả, ấ, ế, ệ, ố, ớ, ...) đã bị thay vĩnh viễn bằng dấu '?' ngay trong file,
còn các chữ như à á â ã è é ê ì í ò ó ô õ ù ú ý vẫn giữ nguyên. Vì vậy KHÔNG thể "convert" ngược về Unicode bằng bảng mã.

Cách xử lý ở đây (theo thứ tự):
1. Nếu phát hiện dữ liệu đúng là TCVN3 thật (nhiều ký tự µ ¸ ¹ ¶ · ...) -> cảnh báo.
2. Khôi phục chữ '?' bằng từ điển: các từ/cụm từ đúng (file tu_dien_phuc_hoi.txt + toàn bộ chữ trong CHECK_2026.xlsx).
   Mỗi cụm đúng được "làm hỏng" theo cùng cách (mã hóa cp1252, ký tự lạ -> '?') để tạo khóa tra cứu.
3. Mã hàng / mã KH là số hoặc ASCII nên phân tích không bị ảnh hưởng.
"""
import os
import re
import unicodedata
from collections import Counter

import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    _ext_dict = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "workflow", "tu_dien_phuc_hoi.txt")
    DICT_FILE = _ext_dict if os.path.exists(_ext_dict) else os.path.join(HERE, "tu_dien_phuc_hoi.txt")
else:
    DICT_FILE = os.path.join(HERE, "tu_dien_phuc_hoi.txt")

TCVN3_HINT = set("µ¸¹¶·¨»¼½¾Æ®ÌÎÏÐÑªÒÓÔÕÖ×ØÜÝÞß")


def lossy(s):
    """Dạng bị hỏng của chuỗi khi lưu cp1252: ký tự ngoài bảng mã -> '?'."""
    return s.encode("cp1252", errors="replace").decode("cp1252").lower()


def looks_like_tcvn3(texts, sample=2000):
    n = bad = 0
    for t in texts:
        if not isinstance(t, str):
            continue
        n += len(t)
        bad += sum(ch in TCVN3_HINT for ch in t)
        if n > sample * 20:
            break
    return n > 0 and bad / n > 0.02


class Restorer:
    def __init__(self, extra_texts=()):
        self.phrases = {}   # khóa hỏng -> cụm đúng (nhiều từ)
        self.words = {}     # khóa hỏng -> Counter các từ đúng
        phrases = []
        if os.path.exists(DICT_FILE):
            for line in open(DICT_FILE, encoding="utf-8"):
                line = line.strip()
                if line and not line.startswith("#"):
                    phrases.append(line)
        for line in phrases:
            self._add(line, boost=1000)
        for t in extra_texts:
            if isinstance(t, str):
                self._add(re.sub(r"\([^)]*\)", " ", t))

    def _add(self, text, boost=1):
        text = unicodedata.normalize("NFC", text)
        toks = re.findall(r"[^\W\d_]+(?:\.[^\W\d_]+)*", text)
        if len(toks) > 1 and boost >= 1000:
            self.phrases[lossy(" ".join(toks))] = " ".join(toks)
        for w in toks:
            if "?" in lossy(w) or lossy(w) != w.lower():
                self.words.setdefault(lossy(w), Counter())[w.lower()] += boost
            else:
                self.words.setdefault(lossy(w), Counter())[w.lower()] += 0
        return toks

    @staticmethod
    def _case_like(orig, new):
        letters = [c for c in orig if c != "?"]
        if letters and letters[0].isupper():
            return new[:1].upper() + new[1:]
        if letters and all(c.isupper() for c in letters) and len(letters) > 1:
            return new.upper()
        return new

    def restore(self, text):
        if not isinstance(text, str) or "?" not in text:
            return text
        out = unicodedata.normalize("NFC", text)
        low = out.lower()
        # 1. cụm từ (dài trước)
        for key in sorted(self.phrases, key=len, reverse=True):
            if "?" not in key:
                continue
            pat = r"(?<![^\W\d_])" + re.escape(key).replace(r"\ ", r"\s+") + r"(?![^\W\d_])"
            for m in list(re.finditer(pat, low))[::-1]:
                rep = self._case_like(out[m.start():m.end()], self.phrases[key])
                out = out[:m.start()] + rep + out[m.end():]
                low = out.lower()
        # 2. từng từ
        def fix(m):
            w = m.group(0)
            if "?" not in w:
                return w
            c = self.words.get(w.lower())
            if not c:
                return w
            best = c.most_common(1)[0][0]
            return self._case_like(w, best)
        return re.sub(r"[^\W\d_?]*\?[^\W\d_]*(?:\?[^\W\d_]*)*", fix, out)


# ======================================================================================
# CHUYỂN MÃ FONT TIẾNG VIỆT CŨ -> UNICODE (NFC)
# ======================================================================================
# Hỗ trợ: TCVN3 (ABC / .Vn*), VNI-Windows, UTF-8 bị đọc nhầm thành cp1252 (mojibake),
# thực thể HTML (&#7879; / &aacute;), Unicode dạng tổ hợp (NFD / CP1258) -> NFC.
# Chưa hỗ trợ: VISCII, VPS, BK-HCM (cần bảng mã riêng, chưa có nguồn đối chiếu đáng tin).
import html as _html

# --- TCVN3 (ABC): byte -> chữ Unicode (byte được lấy qua cp1252 của ký tự đọc ra từ .xls)
_TCVN3_PAIRS = (
    "\xb5à \xb6ả \xb7ã \xb8á \xb9ạ \xa8ă \xbbằ \xbcẳ \xbdẵ \xbeắ \xc6ặ "
    "\xa9â \xc7ầ \xc8ẩ \xc9ẫ \xcaấ \xcbậ \xaeđ "
    "\xccè \xceẻ \xcfẽ \xd0é \xd1ẹ \xaaê \xd2ề \xd3ể \xd4ễ \xd5ế \xd6ệ "
    "\xd7ì \xd8ỉ \xdcĩ \xddí \xdeị "
    "\xdfò \xe1ỏ \xe2õ \xe3ó \xe4ọ \xabô \xe5ồ \xe6ổ \xe7ỗ \xe8ố \xe9ộ "
    "\xacơ \xeaờ \xebở \xecỡ \xedớ \xeeợ "
    "\xefù \xf1ủ \xf2ũ \xf3ú \xf4ụ \xadư \xf5ừ \xf6ử \xf7ữ \xf8ứ \xf9ự "
    "\xfaỳ \xfbỷ \xfcỹ \xfdý \xfeỵ "
    "\xa1Ă \xa2Â \xa3Ê \xa4Ô \xa5Ơ \xa6Ư \xa7Đ"
)
_TCVN3 = {}
for _p in _TCVN3_PAIRS.split():
    _TCVN3[_p[0].encode("latin-1").decode("cp1252", "replace")] = _p[1]
_TCVN3_SIG = set("µ¶·¸¹¨©ª«¬\xad®")   # ký tự đặc trưng của TCVN3 (không có trong chữ Việt cp1252 bị mất dấu)

# --- VNI-Windows: nguyên âm + ký tự dấu (2 ký tự) hoặc 1 ký tự
_VNI_LOWER = {
    "à": "aø", "á": "aù", "ả": "aû", "ã": "aõ", "ạ": "aï",
    "ă": "aê", "ằ": "aè", "ắ": "aé", "ẳ": "aú", "ẵ": "aü", "ặ": "aë",
    "â": "aâ", "ầ": "aà", "ấ": "aá", "ẩ": "aå", "ẫ": "aã", "ậ": "aä",
    "è": "eø", "é": "eù", "ẻ": "eû", "ẽ": "eõ", "ẹ": "eï",
    "ê": "eâ", "ề": "eà", "ế": "eá", "ể": "eå", "ễ": "eã", "ệ": "eä",
    "ì": "ì", "í": "í", "ỉ": "æ", "ĩ": "ó", "ị": "ò",
    "ò": "oø", "ó": "où", "ỏ": "oû", "õ": "oõ", "ọ": "oï",
    "ô": "oâ", "ồ": "oà", "ố": "oá", "ổ": "oå", "ỗ": "oã", "ộ": "oä",
    "ơ": "ô", "ờ": "ôø", "ớ": "ôù", "ở": "ôû", "ỡ": "ôõ", "ợ": "ôï",
    "ù": "uø", "ú": "uù", "ủ": "uû", "ũ": "uõ", "ụ": "uï",
    "ư": "ö", "ừ": "öø", "ứ": "öù", "ử": "öû", "ữ": "öõ", "ự": "öï",
    "ỳ": "yø", "ý": "yù", "ỷ": "yû", "ỹ": "yõ", "ỵ": "î",
    "đ": "ñ",
}
_VNI = {}
for _u, _v in _VNI_LOWER.items():
    _VNI[_v] = _u
    _VNI[_v.upper()] = _u.upper()
_VNI_KEYS = sorted(_VNI, key=len, reverse=True)
_VNI_RE = re.compile("|".join(re.escape(k) for k in _VNI_KEYS))
_VNI_SIG = set("ñöÑÖîÎæÆ")
_MOJIBAKE_RE = re.compile(r"[ÂÃÄÅáâ][\u0080-\u00bf\u0152\u0153\u0160\u0161\u0178\u017d\u017e\u0192\u02c6\u02dc\u2013-\u203a\u20ac\u2122]")


def _nonascii(texts, limit=20000):
    out = []
    for t in texts:
        if isinstance(t, str):
            out.extend(ch for ch in t if ord(ch) > 127)
            if len(out) > limit:
                break
    return out


def detect_encoding(texts):
    """Đoán bảng mã của cả cột văn bản: 'tcvn3' | 'vni' | 'unicode' (đã đúng hoặc chỉ mất dấu '?')."""
    chars = _nonascii(texts)
    if not chars:
        return "unicode"
    n = len(chars)
    if sum(c in _TCVN3_SIG for c in chars) / n >= 0.2:
        return "tcvn3"
    if sum(c in _VNI_SIG for c in chars) >= 3 and sum(c in _VNI_SIG for c in chars) / n >= 0.1:
        return "vni"
    return "unicode"


def convert_tcvn3(s):
    return "".join(_TCVN3.get(ch, ch) for ch in s)


def convert_vni(s):
    return _VNI_RE.sub(lambda m: _VNI[m.group(0)], s)


def to_unicode(s, enc="unicode"):
    """Chuyển 1 chuỗi về Unicode NFC theo bảng mã `enc` (từ detect_encoding)."""
    if not isinstance(s, str) or not s:
        return s
    if "&" in s and re.search(r"&#?\w+;", s):
        s = _html.unescape(s)
    if enc == "tcvn3":
        s = convert_tcvn3(s)
    elif enc == "vni":
        s = convert_vni(s)
    elif _MOJIBAKE_RE.search(s):          # UTF-8 bị đọc thành cp1252
        try:
            s = s.encode("cp1252").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return unicodedata.normalize("NFC", s)
