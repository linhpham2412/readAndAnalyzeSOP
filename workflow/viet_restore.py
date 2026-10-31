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
