# -*- coding: utf-8 -*-
"""
Smart ID Finder v6 - دقت بسیار بالا
- الگوریتم تطبیق سخت‌گیرانه با الزام برند + عدد + کلمات کلیدی
- نمایش Top-N کاندید و انتخاب دستی
- تشخیص قوی‌تر کالاهای غیرمشمول
- وضعیت‌های دقیق‌تر (عالی / خوب / مشکوک / غیرمشمول / موجود نیست)
"""
import sys
import os
import re
import time
import logging
import difflib
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Set
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import requests
from bs4 import BeautifulSoup

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QProgressBar, QSpinBox,
    QGroupBox, QMessageBox, QTextEdit, QHeaderView, QStatusBar, QTabWidget,
    QLineEdit, QDialog, QDialogButtonBox, QFormLayout, QMenu, QAction,
    QListWidget, QListWidgetItem, QAbstractItemView, QCheckBox, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

# ============================================================================
# تنظیمات
# ============================================================================
CONFIG = {
    "default_threshold": 60,          # آستانه بررسی (بر اساس لاگ واقعی)
    "auto_accept_threshold": 78,      # بالای این = خودکار قبول
    "review_threshold": 60,           # بین این و auto = نیاز به بررسی
    "search_delay": 0.25,
    "thdorsan_page": "https://thdorsan.com/%D8%AF%D8%B1%DB%8C%D8%A7%D9%81%D8%AA-%D8%B4%D9%86%D8%A7%D8%B3%D9%87-%DA%A9%D8%A7%D9%84%D8%A7-%D9%88-%D8%AE%D8%AF%D9%85%D8%A7%D8%AA-%D8%B3%D8%A7%D9%85%D8%A7%D9%86%D9%87-%D9%85%D9%88%D8%AF%DB%8C%D8%A7%D9%86",
    "thdorsan_api": "https://thdorsan.com/Stuff",
    "max_candidates_keep": 80,
}

# ─── کالاهایی که در سامانه شناسه کالا معمولاً ثبت نمی‌شوند ───
# کلمات میوه/سبزی — فقط وقتی کالای «تازه» باشد غیرمشمول است
FRESH_PRODUCE_KEYWORDS = {
    'گیلاس', 'آلبالو', 'هلو', 'زردآلو', 'انگور', 'نارنگی', 'لیمو', 'پیاز',
    'گوجه', 'بادمجان', 'تره', 'شلیل', 'راک ملون', 'کلم', 'چاقاله',
    'آلو', 'توت فرنگی', 'خیار', 'کاهو', 'کرفس', 'جعفری', 'ریحان',
    'سیب', 'سیب گلاب', 'سیب زمینی', 'هویج', 'ترب', 'چغندر', 'اسفناج',
    'نعناع', 'شنبلیله', 'انار', 'کیوی', 'موز', 'پرتقال', 'گریپ فروت',
    'هندوانه', 'طالبی', 'خربزه', 'انجیر', 'آووکادو',
}
# اگر این کلمات در نام باشند، میوه/سبزی را «فرآوری‌شده» می‌دانیم (مشمول)
PROCESSED_OVERRIDE = {
    'پودر', 'ژله', 'خشکه', 'خشک', 'کنسرو', 'کمپوت', ' مربا', 'مربا',
    'آبمیوه', 'نکتار', 'پوره', 'چیپس', 'اسنک', 'لواشک', 'برگه',
    'پنیر', 'ماست', 'شیر', 'کره', 'خامه', 'دوغ',
    'شامپو', 'مایع', 'کرم', 'صابون', 'پودر لباسشویی',
}
NON_REGISTERED_KEYWORDS = {
    # موبایل و الکترونیک
    'موبایل', 'موب', 'سامسونگ', 'شیاومی', 'شیائومی', 'اپل', 'آیفون', 'ریکاوری',
    'هواوی', 'نوکیا', 'آنر', 'ریلمی', 'ردمی',
    # پوشاک و لباس زیر
    'دوک', 'شورت', 'جوراب', 'زیرپوش', 'پادار', 'اسلیپ', 'لباس زیر',
    'سوتین', 'مایو', 'تیشرت', 'شلوارک',
    # لوازم تحریر
    'دفتر', 'خودکار', 'مداد', 'پاک کن', 'تراش', 'خط کش', 'اتود',
    'ماژیک', 'لاک', 'غلط گیر', 'منگنه', 'کلاسور', 'پوشه',
    'استیکر', 'هایلایتر', 'علامت زن', 'پرچم', 'قمقمه',
    'جامدادی', 'زیر دستی', 'مقوا', 'مداد رنگی', 'پاستل', 'گواش', 'آبرنگ',
    # متفرقه
    'لیبل', 'کیسه توری', 'میوه لوکس', 'میوه کی اف سی',
    'بستنی', 'زغال', 'یخ', 'یخ خشک',
}
# انواع کالای مهم برای تقویت تطبیق
PRODUCT_TYPE_WORDS = {
    'شامپو', 'بدن', 'مام', 'اسپری', 'کرم', 'آبرسان', 'مرطوب',
    'پنیر', 'کره', 'شیر', 'ماست', 'خامه', 'دوغ',
    'کیک', 'بیسکویت', 'کلوچه', 'کراکر', 'اسنک', 'چیپس', 'پف',
    'نوار', 'دستمال', 'پوشک',
    'مایع', 'دستشویی', 'لباسشویی', 'ظرفشویی',
    'کنسرو', 'رب', 'سس', 'روغن',
    'آبمیوه', 'نوشابه', 'دوغ',
}

# ─── کلمات حذفی (Stop Words) ───
STOP_WORDS = {
    'گرمی', 'گرم', 'عددی', 'عدد', 'سی', 'میلی', 'میل', 'لیتری', 'لیتر',
    'بزرگ', 'کوچک', 'متوسط', 'خانواده', 'مجلسی', 'اقتصادی', 'ویژه',
    'ساده', 'تک', 'پک', 'بسته', 'رول', 'در', 'برگ', 'g', 'x',
    'کیلویی', 'کیلو', 'جعبه', 'شیشه', 'بطری', 'پت', 'پاکتی',
    'قوطی', 'کنسروی', 'پریمیوم', 'لوکس', 'معمولی', 'کلاسیک', 'فانتزی',
    'ع', 'ل', 'م', 'گ', 'میلی لیتر', 'میلی‌لیتر', 'میلی لیتری',
    'عدد ی', 'عددی', 'بسته ای', 'بسته\u200cای', 'رولی',
    'ای', 'لایه', 'مردانه', 'زنانه', 'کودک', 'نوزاد',
    # کلمات رایج در شرح رسمی سامانه که ارزش تطبیق ندارند
    'نام', 'تجارتی', 'برند', 'سازنده', 'تولید', 'کننده', 'وارداتی',
    'اختصاصی', 'عمومی', 'طعم', 'نوع', 'مدل', 'کد', 'شناسه',
    'سیل', 'اسپرید', 'شماره',
}

# ─── برندهای چند کلمه‌ای (مرتب‌شده بر اساس طول برای اولویت) ───
MULTI_BRANDS_RAW = [
    'چی توز', 'شیرین عسل', 'کاله', 'پگاه', 'مزمز', 'نادری', 'نیک شف',
    'کیمبال', 'بیوگلز', 'سمین', 'چکاد', 'هانیپرز', 'سوربن', 'ال پی', 'الپی',
    'بی اند کی', 'بی اندکی', 'کامان', 'دیپ سنس', 'شیبا', 'فرمند', 'گلها',
    'فیروز', 'اوه', 'شون', 'مای', 'کالین', 'فیبورا', 'راشین', 'بادوک',
    'اچ سی', 'اج سی', 'سیرنگ', 'پیکولا', 'لوتوس', 'ترش اکس', 'ویت یو',
    'بیول', 'ریجوی', 'ویتالیر', 'درماتیپیک', 'کالیستا', 'کاپیتانو',
    'بن مانو', 'های فود', 'کیالو', 'مهیا', 'بیتا', 'آلیما', 'جلفا',
    'بیستون', 'شارین', 'ناتسور', 'رنکس', 'اکشن', 'دکتر بن', 'پنتر',
    'کلیپس', 'تی تی ام', 'فابل', 'هیپو', 'کلاس ممو', 'ممو استایل',
    'بیرو', 'آبی', 'بیک', 'کنکو', 'اطلس', 'سی کلاس', 'لیف پن', 'دلی',
    'پالمو', 'اینوکس', 'آبنوس', 'صدف', 'پرشیا', 'شاپیکان', 'شاپیگان',
    'کریستال', 'گاریمپا', 'پاپیا', 'تست', 'مولپد', 'ناتار', 'سیف',
    'دامستوس', 'امومایع', 'کامفورت', 'آکس', 'پینو', 'لندی', 'پدکس',
    'دارکوب', 'پیکنیک', 'پیک نیک', 'برکت', 'اورمو', 'کایسر', 'کاتوس',
    'اسکول مکس', 'ماکان', 'بارمان', 'پارسیکار', 'اسکینشیک', 'لروکس',
    'ژبن پ ب', 'ای آی ان', 'زاگو', 'کیوب لند', 'کاسل', 'تی ان تی',
    'دیگنیتی', 'هوم هیرو', 'ماهور', 'کیمدی', 'آذوقه', 'نازدوخت',
    'گرین فیلد', 'گلداستار', 'نظری', 'یسی مانکیی', 'نیچرلین', 'سیلور',
    'کوهپناه', 'کیان', 'سن ایچ', 'فروتیوه', 'افس', 'کانادادرای',
    'جوجو', 'داینامین', 'لاکی دو', 'هی دی', 'بهوش', 'کولاک', 'دسانی',
    'بهنوش', 'پاک', 'هراز', 'صباح', 'سمیه', 'رضوی',
    'سولیکو', 'فارسی', 'تانیس', 'اورادنت', 'مریدنت', 'میلکی', 'گودمود',
    'پنبه ریز', 'گل پرنوار', 'بی تا', 'مای لیدی', 'مای بی بی',
    'ایزی لایف', 'نینو', 'ثمین', 'اوک شاین', 'نیوتیس', 'سوپکس', 'شاوما',
    'ب.آ', 'سحرسان', 'وارمی', 'الارو', 'سی گل', 'پرایم', 'آوات',
    'آریا', 'اریا', 'توتو', 'ممو', 'کلاس', 'فکتیس', 'ام جی ام',
    'ام پی', 'کرون', 'مکسنس', 'سلنا', 'کی دبلیو تریو',
    'کینوفیت', 'پاپس', 'آیسول', 'سافتلن', 'یونی',
    'اسپرینگ', 'میهن', 'دامداران', 'کاله پرو', 'کاله پرون',
    'های پک', 'هایپک', 'دلپذیر', 'رشد', 'زر ماکارون', 'زر',
    'هاتی کارا', 'هاتیکارا', 'مهرام', 'یک و یک', 'یک‌ویک',
    'بیز', 'بیزکویین', 'بیسکویت', 'نان رضوی',
]

# مرتب‌سازی بر اساس طول (بلندتر اول) برای تطابق بهتر
MULTI_BRANDS = sorted(set(MULTI_BRANDS_RAW), key=lambda x: len(x), reverse=True)


# ============================================================================
# ابزارهای کمکی
# ============================================================================
def normalize(text: str) -> str:
    """نرمال‌سازی متن فارسی/انگلیسی"""
    if not text:
        return ""
    text = str(text).lower()
    # حذف علائم نگارشی رایج فارسی و انگلیسی (ویرگول فارسی، نقطه و ...)
    text = re.sub(r'[،,؛;:؟!؟\.…\-–—_/\\|()\[\]{}«»""\'\"+*=]', ' ', text)
    # حذف هر کاراکتر غیرحرفی/غیرعددی/غیر فاصله (با حفظ حروف فارسی)
    text = re.sub(r'[^\w\s\u0600-\u06FF]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    # ارقام فارسی → انگلیسی
    for p, e in {
        '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
        '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
    }.items():
        text = text.replace(p, e)
    # یکسان‌سازی ی و ک
    text = text.replace('ي', 'ی').replace('ك', 'ک')
    # نیم‌فاصله و فاصله‌های خاص
    text = text.replace('\u200c', ' ').replace('\u200f', ' ').replace('\xa0', ' ')
    return text


def extract_numbers(text: str) -> Dict[str, float]:
    """استخراج اعداد با واحد (وزن/حجم/تعداد) — پشتیبانی از فاصله و بدون فاصله"""
    text = str(text)
    for p, e in {
        '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
        '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
    }.items():
        text = text.replace(p, e)

    numbers: Dict[str, float] = {}
    patterns = [
        (r'(\d+(?:[./]\d+)?)\s*(?:گرمی|گرم|گ)\b', 'گرم'),
        (r'(\d+(?:[./]\d+)?)(?:گرمی|گرم)\b', 'گرم'),
        (r'(\d+(?:[./]\d+)?)\s*(?:میلی\s*لیتر|میلی‌لیتر|میلی|میل|مل)\b', 'میل'),
        (r'(\d+(?:[./]\d+)?)(?:میلی|میل|مل)\b', 'میل'),
        (r'(\d+(?:[./]\d+)?)\s*(?:لیتری|لیتر)\b', 'لیتر'),
        (r'(\d+(?:[./]\d+)?)(?:لیتری|لیتر)\b', 'لیتر'),
        (r'(\d+(?:[./]\d+)?)\s*(?:سی)\b', 'سی'),
        (r'(\d+(?:[./]\d+)?)\s*(?:عددی|عدد)\b', 'عدد'),
        (r'(\d+(?:[./]\d+)?)\s*(?:برگ)\b', 'برگ'),
        (r'(\d+(?:[./]\d+)?)\s*(?:رول)\b', 'رول'),
        (r'(\d+(?:[./]\d+)?)\s*(?:کیلویی|کیلو)\b', 'کیلو'),
        (r'(\d+(?:[./]\d+)?)\s*(?:میلی\s*گرم|میلی‌گرم)\b', 'میلی‌گرم'),
    ]
    for pattern, unit in patterns:
        for m in re.findall(pattern, text, flags=re.IGNORECASE):
            try:
                raw = m[0] if isinstance(m, tuple) else m
                val = float(str(raw).replace('/', '.'))
                if unit not in numbers:
                    numbers[unit] = val
            except Exception:
                pass
    return numbers


def extract_keywords(text: str, brand: str = "") -> Set[str]:
    """استخراج کلمات کلیدی معنادار (بدون عدد، stop-word و برند)"""
    text = normalize(text)
    # حذف برند کامل از متن تا تکه‌تکه نشود
    if brand:
        text = text.replace(brand, ' ')
    text = re.sub(r'\d+', ' ', text)
    words = text.split()
    kws = {w for w in words if w and len(w) >= 2 and w not in STOP_WORDS}
    # اگر برند چندکلمه‌ای بود، تکه‌هایش را هم حذف کن
    if brand:
        for part in brand.split():
            kws.discard(part)
    return kws


def extract_brand(text: str) -> str:
    """استخراج برند از نام کالا (اولویت برندهای چندکلمه‌ای)"""
    norm = normalize(text)
    for brand in MULTI_BRANDS:
        if brand in norm:
            return brand
    words = [w for w in norm.split() if w and len(w) >= 2 and w not in STOP_WORDS]
    return words[0] if words else ""


def is_non_registered(text: str) -> bool:
    """
    بررسی آیا کالا در سامانه ثبت نمی‌شود.
    میوه/سبزی تازه → غیرمشمول
    میوه در محصول فرآوری‌شده (پودر ژله طالبی، آلبالو خشکه) → مشمول
    """
    norm = normalize(text)

    # استثنای فرآوری‌شده: اگر کلمه فرآوری باشد، میوه را غیرمشمول نکن
    is_processed = any(p.strip() in norm for p in PROCESSED_OVERRIDE)

    for kw in NON_REGISTERED_KEYWORDS:
        if kw in norm:
            return True

    # میوه و سبزی تازه
    for kw in FRESH_PRODUCE_KEYWORDS:
        if kw in norm:
            if is_processed:
                continue  # مثلاً پودر ژله طالبی
            # اگر نام تقریباً فقط همان میوه است (کوتاه)
            words = [w for w in norm.split() if len(w) >= 2]
            if len(words) <= 3 or kw == norm or norm.startswith(kw):
                return True
            # «سیب گلاب»، «انگور قرمز» و مشابه
            if any(w in FRESH_PRODUCE_KEYWORDS or w in ('قرمز', 'سبز', 'زرد', 'سفید', 'گلاب', 'بیدانه', 'پاکستانی', 'گلخانه', 'فرنگی') for w in words):
                return True
    return False


def token_set_ratio(a: str, b: str) -> float:
    """شباهت مبتنی بر مجموعه توکن‌ها (شبیه fuzzywuzzy token_set_ratio)"""
    ta = extract_keywords(a)
    tb = extract_keywords(b)
    if not ta or not tb:
        return 0.0
    inter = ta & tb
    if not inter:
        return 0.0
    # نسبت اشتراک به کوچک‌ترین مجموعه + پاداش برای overlap خوب
    ratio = len(inter) / min(len(ta), len(tb))
    return min(100.0, ratio * 100)


# ============================================================================
# سیستم امتیازدهی پیشرفته و سخت‌گیرانه (v6)
# ============================================================================
def compute_score(target: str, candidate: str) -> Dict:
    """
    محاسبه امتیاز جامع و سخت‌گیرانه.
    قوانین کلیدی برای دقت بالا:
      - اگر برند استخراج‌شده در نام کاندید نباشد → جریمه سنگین
      - اگر عدد وزن/حجم وجود داشته باشد و مطابقت نکند → جریمه سنگین
      - کلمات کلیدی مشترک وزن بالایی دارند
    """
    t_norm = normalize(target)
    c_norm = normalize(candidate)

    if not t_norm or not c_norm:
        return {
            'final': 0, 'ratio': 0, 'partial': 0, 'keyword': 0,
            'number': 0, 'brand_match': 0, 'token_set': 0,
            'details': 'خالی', 'pass_brand': False, 'pass_number': True
        }

    # 1. SequenceMatcher کامل
    ratio = difflib.SequenceMatcher(None, t_norm, c_norm).ratio() * 100

    # 2. Partial (زیررشته)
    shorter = t_norm if len(t_norm) <= len(c_norm) else c_norm
    longer = c_norm if len(t_norm) <= len(c_norm) else t_norm
    partial = 0.0
    if shorter:
        for i in range(max(1, len(longer) - len(shorter) + 1)):
            s = difflib.SequenceMatcher(None, shorter, longer[i:i + len(shorter)]).ratio()
            if s > partial:
                partial = s
    partial *= 100

    # برند را زودتر استخراج می‌کنیم تا از keywords حذف شود
    brand_t = extract_brand(target)

    # 3. کلمات کلیدی (Jaccard-like) — برند جدا شده
    kw_t = extract_keywords(target, brand_t)
    kw_c = extract_keywords(candidate, brand_t)
    if kw_t and kw_c:
        intersection = kw_t & kw_c
        union = kw_t | kw_c
        keyword = (len(intersection) / len(union)) * 100 if union else 0
        # پاداش برای پوشش خوب نسبت به هدف
        coverage = (len(intersection) / len(kw_t)) * 100 if kw_t else 0
        keyword = max(keyword, coverage * 0.85)
    else:
        keyword = 0.0

    # 4. Token-set (مشابه fuzzy) — روی کل متن
    token_set = token_set_ratio(target, candidate)

    # 5. تطابق اعداد
    # نکته مهم از لاگ واقعی: خیلی از شرح‌های رسمی وزن ندارند.
    # پس نبود عدد در کاندید = خنثی (نه جریمه سنگین).
    nums_t = extract_numbers(target)
    nums_c = extract_numbers(candidate)
    number_score = 50.0
    pass_number = True
    number_mismatch = False

    if nums_t and nums_c:
        common_units = set(nums_t.keys()) & set(nums_c.keys())
        if common_units:
            matches = 0
            for u in common_units:
                vt, vc = nums_t[u], nums_c[u]
                if abs(vt - vc) / max(vt, vc, 0.01) <= 0.15 or abs(vt - vc) <= 2:
                    matches += 1
            number_score = (matches / len(common_units)) * 100
            if number_score < 50:
                number_mismatch = True
                pass_number = False
            else:
                pass_number = True
        else:
            # واحد متفاوت — خنثی ملایم
            number_score = 40.0
    # اگر فقط هدف عدد دارد و کاندید ندارد → خنثی بمان (50)

    # 6. تطابق برند
    brand_match = 0.0
    pass_brand = True
    if brand_t:
        if brand_t in c_norm:
            brand_match = 100.0
            pass_brand = True
        else:
            brand_ok = False
            for w in c_norm.split():
                if brand_t in w or w in brand_t:
                    if len(w) >= 3 and difflib.SequenceMatcher(None, brand_t, w).ratio() > 0.75:
                        brand_ok = True
                        break
            if brand_ok:
                brand_match = 75.0
                pass_brand = True
            else:
                brand_match = 0.0
                pass_brand = False
    else:
        brand_match = 40.0
        pass_brand = True

    # 7. پاداش نوع کالا (شامپو، پنیر، کره، ...)
    type_t = {w for w in extract_keywords(target, brand_t) if w in PRODUCT_TYPE_WORDS}
    type_c = {w for w in normalize(candidate).split() if w in PRODUCT_TYPE_WORDS}
    type_boost = 0.0
    if type_t and type_c and (type_t & type_c):
        type_boost = 12.0  # پاداش ثابت برای هم‌نوع بودن

    # ─── محاسبه نهایی (نسخه متعادل بر اساس لاگ واقعی) ───
    # برند مهم‌ترین سیگنال است؛ عدد فقط وقتی هر دو طرف دارند سخت‌گیری می‌شود.
    final = (
        brand_match * 0.32 +
        (keyword * 0.20 + token_set * 0.15) +
        number_score * 0.12 +
        (ratio * 0.06 + partial * 0.08) +
        type_boost
    )

    # جریمه فقط وقتی برند قطعاً اشتباه است
    if brand_t and not pass_brand:
        final *= 0.40

    # جریمه عدد فقط وقتی هر دو طرف عدد دارند و با هم نمی‌خوانند
    if number_mismatch:
        final *= 0.75

    # اگر برند درست است ولی هیچ کلمه کلیدی مشترکی نیست، کمی کم کن نه نابود
    if brand_t and pass_brand and kw_t and keyword < 15 and token_set < 20:
        final *= 0.85

    # کف امتیاز وقتی برند + نوع کالا هر دو درست‌اند
    if brand_t and pass_brand and type_boost > 0:
        final = max(final, 68.0)

    final = max(0.0, min(final, 100.0))

    return {
        'final': final,
        'ratio': ratio,
        'partial': partial,
        'keyword': keyword,
        'number': number_score,
        'brand_match': brand_match,
        'token_set': token_set,
        'pass_brand': pass_brand,
        'pass_number': pass_number,
        'details': (
            f"R:{ratio:.0f} P:{partial:.0f} K:{keyword:.0f} "
            f"T:{token_set:.0f} N:{number_score:.0f} B:{brand_match:.0f}"
        )
    }


# ============================================================================
# موتور جستجو
# ============================================================================
class SearchEngine:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'X-Requested-With': 'XMLHttpRequest',
            'Referer': CONFIG['thdorsan_page'],
            'Accept-Language': 'fa-IR,fa;q=0.9,en;q=0.8',
        })
        self.token = None
        self.cache: Dict[str, List[Dict]] = {}

    def init_session(self, log_cb=None) -> bool:
        if self.token:
            return True
        try:
            r = self.session.get(CONFIG['thdorsan_page'], timeout=18)
            soup = BeautifulSoup(r.text, 'html.parser')
            meta = soup.find('meta', {'name': 'csrf-token'})
            if meta and meta.get('content'):
                self.token = meta['content']
            else:
                m = re.search(r"_token\s*[:=]\s*['\"]([a-zA-Z0-9]+)['\"]", r.text)
                if m:
                    self.token = m.group(1)
            if self.token:
                self.session.headers['X-CSRF-TOKEN'] = self.token
                if log_cb:
                    log_cb("✅ توکن CSRF دریافت شد")
                return True
            if log_cb:
                log_cb("⚠️ توکن CSRF پیدا نشد")
        except Exception as e:
            if log_cb:
                log_cb(f"⚠️ خطا در اتصال: {e}")
        return False

    def _api_search(self, keyword: str) -> List[Dict]:
        if not keyword or len(keyword.strip()) < 2:
            return []
        try:
            r = self.session.post(
                CONFIG['thdorsan_api'],
                data={'keyword': keyword.strip()},
                timeout=18
            )
            if r.status_code == 200:
                data = r.json().get('data', [])
                return [{
                    'id': str(i.get('id', '')),
                    'name': i.get('description', '') or '',
                    'type': i.get('type', ''),
                    'vat': str(i.get('vat', '')),
                } for i in data if i.get('id')]
        except Exception:
            pass
        return []

    def generate_search_terms(self, product_name: str, manual_kw: str = "") -> List[str]:
        """
        تولید هوشمند کلمات جستجو با اولویت:
        1. کلمات دستی کاربر
        2. برند + کلمه کلیدی اصلی
        3. فقط برند
        4. کلمات کلیدی مهم
        """
        terms: List[str] = []

        # 1. کلمات دستی
        if manual_kw and manual_kw.strip():
            terms.append(manual_kw.strip())
            for part in re.split(r'[\s,،]+', manual_kw.strip()):
                part = part.strip()
                if len(part) >= 3 and part not in terms:
                    terms.append(part)

        brand = extract_brand(product_name)
        keywords = list(extract_keywords(product_name, brand))
        # اولویت به کلمات نوع کالا
        keywords.sort(key=lambda w: (0 if w in PRODUCT_TYPE_WORDS else 1, -len(w)))

        # 2. برند + کلمات کلیدی
        if brand and len(brand) >= 2:
            if keywords:
                terms.append(f"{brand} {keywords[0]}")
                if len(keywords) > 1:
                    terms.append(f"{brand} {keywords[1]}")
                if len(keywords) >= 2 and len(keywords[0]) + len(keywords[1]) < 20:
                    terms.append(f"{brand} {keywords[0]} {keywords[1]}")
            terms.append(brand)

        # 3. کلمات کلیدی معنادار (حداقل ۳ حرف، نه stop)
        for kw in keywords[:3]:
            if len(kw) >= 3 and kw not in terms:
                terms.append(kw)

        # حذف تکرار، کلمات خیلی کوتاه و نویز
        noise = {'میل', 'ای', 'گ', 'م', 'ل', 'ع', 'سی', 'جی'}
        result = []
        seen = set()
        for t in terms:
            t_clean = t.strip()
            if not t_clean or t_clean in seen:
                continue
            if t_clean in noise or len(t_clean) < 2:
                continue
            # تک‌کلمه زیر ۳ حرف را رد کن (مگر برند)
            parts = t_clean.split()
            if len(parts) == 1 and len(t_clean) < 3 and t_clean != brand:
                continue
            seen.add(t_clean)
            result.append(t_clean)
        return result

    def search(self, product_name: str, manual_kw: str = "",
               log_cb=None) -> Tuple[List[Dict], List[str]]:
        cache_key = f"{product_name}||{manual_kw}"
        if cache_key in self.cache:
            return self.cache[cache_key], ["کش"]

        if is_non_registered(product_name):
            if log_cb:
                log_cb("  ⛔ این نوع کالا در سامانه ثبت نمی‌شود")
            return [], ["غیرمشمول"]

        if not self.token and not self.init_session(log_cb):
            return [], ["خطا در اتصال"]

        terms = self.generate_search_terms(product_name, manual_kw)
        if log_cb:
            log_cb(f"  📝 کلمات جستجو: {terms}")

        all_results: List[Dict] = []
        search_log: List[str] = []

        for term in terms:
            if log_cb:
                log_cb(f"  🔍 '{term}'...")
            results = self._api_search(term)
            if results:
                search_log.append(f"'{term}': {len(results)}")
                all_results.extend(results)
                if len(all_results) >= CONFIG['max_candidates_keep']:
                    break
            time.sleep(0.05)  # کمی فاصله بین درخواست‌ها

        # حذف تکرار بر اساس id
        seen = set()
        unique = []
        for r in all_results:
            if r['id'] not in seen:
                seen.add(r['id'])
                unique.append(r)

        self.cache[cache_key] = unique
        return unique, search_log


# ============================================================================
# موتور تطبیق (سخت‌گیرانه)
# ============================================================================
class Matcher:
    def __init__(self, threshold: float = 60, auto_accept: float = 78):
        self.threshold = threshold
        self.auto_accept = auto_accept

    def find_top_n(self, product_name: str, candidates: List[Dict],
                   n: int = 5) -> List[Tuple[Dict, Dict]]:
        scored = []
        for c in candidates:
            scores = compute_score(product_name, c['name'])
            scored.append((c, scores))
        scored.sort(key=lambda x: x[1]['final'], reverse=True)
        return scored[:n]

    def decide(self, product_name: str, candidates: List[Dict]
               ) -> Tuple[Optional[Dict], Dict, str, List[Tuple[Dict, Dict]]]:
        """
        تصمیم‌گیری نهایی:
        - auto: شناسه را قبول کن
        - review: نیاز به بررسی دستی (امتیاز متوسط)
        - reject: رد کن
        برمی‌گرداند: (best_candidate, score_dict, decision, top_list)
        """
        if not candidates:
            return None, {'final': 0}, 'reject', []

        top = self.find_top_n(product_name, candidates, n=5)
        best, best_score = top[0]
        sim = best_score['final']

        # قوانین سخت‌گیرانه
        if not best_score.get('pass_brand', True) and extract_brand(product_name):
            # برند مطابقت ندارد → معمولاً رد
            if sim < 75:
                return best, best_score, 'reject', top

        if sim >= self.auto_accept and best_score.get('pass_brand', True):
            return best, best_score, 'auto', top
        if sim >= self.threshold:
            return best, best_score, 'review', top
        return best, best_score, 'reject', top


# ============================================================================
# پردازشگر اکسل
# ============================================================================
class ExcelHandler:
    @staticmethod
    def read_excel(fp: str):
        wb = load_workbook(fp, data_only=True)
        ws = wb.active
        headers, rows, kc = [], [], {}

        for ri, row in enumerate(ws.iter_rows(values_only=True)):
            rd = [str(c).strip() if c is not None else "" for c in row]

            if not headers and any('Store' in str(c) for c in rd):
                headers = rd
                for i, c in enumerate(headers):
                    cs = str(c).strip()
                    if cs == 'Store':
                        kc['store'] = i
                    elif cs == 'Barcode':
                        kc['barcode'] = i
                    elif cs == 'InnerShenasehCode':
                        kc['shenaseh'] = i
                    elif cs == 'ItemCode':
                        kc['item_code'] = i
                    elif cs == 'LongDescription':
                        kc['description'] = i
                continue

            # هدرهای تکراری را رد کن
            if headers and any('Store' in str(c) for c in rd):
                continue
            if all(not c for c in rd):
                continue
            rows.append(rd)

        return rows, headers, kc

    @staticmethod
    def save_excel(fp: str, headers: list, rows: list, results: list, kc: dict) -> str:
        wb = Workbook()
        ws = wb.active
        ws.title = "نتایج"

        nh = list(headers) + ["شناسه پیدا شده", "نام تطبیق یافته", "درصد", "وضعیت", "جزئیات امتیاز"]
        ws.append(nh)

        hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hfont = Font(color="FFFFFF", bold=True)
        for ci in range(1, len(nh) + 1):
            c = ws.cell(row=1, column=ci)
            c.fill = hfill
            c.font = hfont
            c.alignment = Alignment(horizontal='center')

        si = kc.get('shenaseh', -1)

        for idx, row in enumerate(rows):
            if idx < len(results):
                r = results[idx]
                rc = list(row)
                if si >= 0 and si < len(rc) and r.get('found_id'):
                    rc[si] = r['found_id']
                nr = rc + [
                    r.get('found_id', ''),
                    r.get('found_name', ''),
                    f"{r.get('similarity', 0):.1f}%",
                    r.get('status', ''),
                    r.get('score_details', '')
                ]
            else:
                nr = list(row) + ["", "", "", "", ""]
            ws.append(nr)

            if idx < len(results):
                sim = results[idx].get('similarity', 0)
                st = results[idx].get('status', '')
                sc = ws.cell(row=idx + 2, column=len(nh) - 1)
                if '✅' in st or sim >= 82:
                    sc.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                elif '🔶' in st or '🔍' in st or (55 <= sim < 82):
                    sc.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                else:
                    sc.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        for ci in range(1, len(nh) + 1):
            ml = max(
                (len(str(c.value)) for row in ws.iter_rows(min_col=ci, max_col=ci)
                 for c in row if c.value),
                default=8
            )
            ws.column_dimensions[get_column_letter(ci)].width = min(ml + 2, 55)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = f"{os.path.splitext(fp)[0]}_results_{ts}.xlsx"
        wb.save(out)
        return out


# ============================================================================
# دیالوگ انتخاب از بین کاندیدها
# ============================================================================
class CandidateSelectDialog(QDialog):
    """نمایش Top-N کاندید و امکان انتخاب دستی"""
    def __init__(self, product_name: str, top_list: List[Tuple[Dict, Dict]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔍 انتخاب شناسه از کاندیدها")
        self.setMinimumSize(780, 420)
        self.selected = None  # (id, name, score)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"<b>کالا:</b> {product_name}"))
        layout.addWidget(QLabel("کاندیدهای برتر را انتخاب کنید (یا رد کنید):"))

        self.listw = QListWidget()
        self.listw.setSelectionMode(QAbstractItemView.SingleSelection)
        for cand, scores in top_list:
            sim = scores['final']
            text = (
                f"[{sim:.1f}%]  {cand['id']}  |  {cand['name'][:90]}\n"
                f"     {scores['details']}"
            )
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, (cand['id'], cand['name'], sim, scores.get('details', '')))
            if sim >= 82:
                item.setBackground(QColor(200, 255, 200))
            elif sim >= 65:
                item.setBackground(QColor(255, 255, 200))
            else:
                item.setBackground(QColor(255, 220, 220))
            self.listw.addItem(item)
        layout.addWidget(self.listw)

        btns = QHBoxLayout()
        ok_btn = QPushButton("✅ انتخاب این شناسه")
        ok_btn.clicked.connect(self._accept_selected)
        skip_btn = QPushButton("❌ رد / بدون شناسه")
        skip_btn.clicked.connect(self.reject)
        btns.addWidget(ok_btn)
        btns.addWidget(skip_btn)
        layout.addLayout(btns)

        if self.listw.count() > 0:
            self.listw.setCurrentRow(0)

    def _accept_selected(self):
        item = self.listw.currentItem()
        if item:
            self.selected = item.data(Qt.UserRole)
            self.accept()
        else:
            QMessageBox.warning(self, "هشدار", "یک مورد را انتخاب کنید.")


class ManualEditDialog(QDialog):
    def __init__(self, product_name, current_kw="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("✏️ ویرایش کلمات جستجو")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"محصول: {product_name}"))
        form = QFormLayout()
        self.kw_edit = QLineEdit(current_kw)
        self.kw_edit.setPlaceholderText("کلمات کلیدی دقیق‌تر (مثال: کره پاستوریزه کاله ۵۰ گرم)")
        form.addRow("کلمات جستجو:", self.kw_edit)
        layout.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_keywords(self):
        return self.kw_edit.text().strip()


# ============================================================================
# Thread پردازش
# ============================================================================
class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)
    result_ready = pyqtSignal(list)
    error_occurred = pyqtSignal(str)
    log_message = pyqtSignal(str)
    # سیگنال برای درخواست بررسی دستی (در حالت interactive)
    need_review = pyqtSignal(int, str, list)  # row_idx, product_name, top_list

    def __init__(self, rows, kc, threshold, auto_accept, engine, matcher,
                 manual_map, interactive_review=False):
        super().__init__()
        self.rows = rows
        self.kc = kc
        self.threshold = threshold
        self.auto_accept = auto_accept
        self.engine = engine
        self.matcher = matcher
        self.matcher.threshold = threshold
        self.matcher.auto_accept = auto_accept
        self.manual_map = manual_map
        self.interactive_review = interactive_review
        self._paused = False
        self._stopped = False
        self._review_result = None  # برای دریافت نتیجه دیالوگ

    def pause(self):
        self._paused = True

    def resume(self):
        self._paused = False

    def stop(self):
        self._stopped = True
        self._paused = False

    @property
    def is_paused(self):
        return self._paused

    def set_review_result(self, result):
        """نتیجه انتخاب کاربر از دیالوگ (id, name, sim, details) یا None"""
        self._review_result = result

    def run(self):
        results = []
        try:
            total = len(self.rows)
            di = self.kc.get('description', -1)
            if di == -1:
                self.error_occurred.emit("ستون LongDescription پیدا نشد!")
                return

            self.log_message.emit(f"🚀 شروع پردازش {total} کالا...")
            self.log_message.emit(
                f"   آستانه بررسی: {self.threshold}% | قبول خودکار: {self.auto_accept}%"
            )
            if not self.engine.init_session(self.log_message.emit):
                self.error_occurred.emit("خطا در اتصال به سایت درسان!")
                return

            found = 0

            for idx, row in enumerate(self.rows):
                if self._stopped:
                    self.log_message.emit(f"⏹️ توقف در ردیف {idx + 1}")
                    break

                while self._paused and not self._stopped:
                    time.sleep(0.12)

                if self._stopped:
                    self.log_message.emit(f"⏹️ توقف در ردیف {idx + 1}")
                    break

                pn = row[di] if di < len(row) else ""
                mkw = self.manual_map.get(idx, "")

                if not pn or not pn.strip():
                    results.append({
                        'found_id': '', 'found_name': '', 'similarity': 0,
                        'status': 'بدون نام', 'search_info': '',
                        'score_details': '', 'candidates': []
                    })
                    continue

                prog = int((idx + 1) / total * 100)
                self.progress_updated.emit(prog, f"[{idx + 1}/{total}] {pn[:45]}...")
                self.log_message.emit(f"\n🔍 [{idx + 1}/{total}] {pn}")
                if mkw:
                    self.log_message.emit(f"  📝 دستی: {mkw}")

                try:
                    if is_non_registered(pn):
                        self.log_message.emit("  ⛔ غیرمشمول - در سامانه ثبت نمی‌شود")
                        results.append({
                            'found_id': '', 'found_name': '', 'similarity': 0,
                            'status': '⛔ غیرمشمول (در سامانه نیست)',
                            'search_info': 'غیرمشمول',
                            'score_details': '', 'candidates': []
                        })
                        continue

                    cands, slog = self.engine.search(pn, mkw, self.log_message.emit)

                    if not cands:
                        self.log_message.emit("  ❌ در سایت موجود نیست")
                        results.append({
                            'found_id': '', 'found_name': '', 'similarity': 0,
                            'status': '❌ در سایت موجود نیست',
                            'search_info': ' | '.join(slog),
                            'score_details': '', 'candidates': []
                        })
                        continue

                    self.log_message.emit(f"  📊 {len(cands)} نتیجه خام")

                    best, scores, decision, top = self.matcher.decide(pn, cands)
                    sim = scores.get('final', 0)

                    # ذخیره کاندیدها برای استفاده بعدی در UI
                    cand_info = [
                        {'id': c['id'], 'name': c['name'], 'score': s['final'],
                         'details': s.get('details', '')}
                        for c, s in top
                    ]

                    if decision == 'auto':
                        st = "✅ تطابق عالی"
                        results.append({
                            'found_id': best['id'],
                            'found_name': best['name'],
                            'similarity': sim,
                            'status': st,
                            'search_info': ' | '.join(slog),
                            'score_details': scores.get('details', ''),
                            'candidates': cand_info
                        })
                        found += 1
                        self.log_message.emit(
                            f"  {st}: {best['name'][:55]} ({sim:.1f}%) [{scores.get('details')}]"
                        )

                    elif decision == 'review':
                        # در حالت interactive می‌توان دیالوگ نشان داد؛
                        # فعلاً شناسه را نمی‌گذاریم و وضعیت «نیاز به بررسی» می‌دهیم
                        st = "🔍 نیاز به بررسی دستی"
                        results.append({
                            'found_id': '',  # عمداً خالی تا کاربر تأیید کند
                            'found_name': best['name'] if best else '',
                            'similarity': sim,
                            'status': st,
                            'search_info': ' | '.join(slog),
                            'score_details': scores.get('details', ''),
                            'candidates': cand_info,
                            '_best_id': best['id'] if best else '',
                            '_best_name': best['name'] if best else ''
                        })
                        self.log_message.emit(
                            f"  {st}: بهترین کاندید {sim:.1f}% — شناسه جایگذاری نشد (بررسی کنید)"
                        )

                    else:  # reject
                        st = "❌ تطابق ضعیف / رد شده"
                        results.append({
                            'found_id': '',
                            'found_name': best['name'] if best else '',
                            'similarity': sim,
                            'status': st,
                            'search_info': ' | '.join(slog),
                            'score_details': scores.get('details', ''),
                            'candidates': cand_info
                        })
                        if best:
                            self.log_message.emit(
                                f"  {st} ({sim:.1f}%) — {scores.get('details')}"
                            )
                        else:
                            self.log_message.emit("  ❌ بدون تطابق")

                except Exception as e:
                    self.log_message.emit(f"  ❌ خطا: {str(e)[:80]}")
                    results.append({
                        'found_id': '', 'found_name': '', 'similarity': 0,
                        'status': 'خطا', 'search_info': '',
                        'score_details': '', 'candidates': []
                    })

                time.sleep(CONFIG['search_delay'])

            self.log_message.emit(f"\n{'=' * 55}")
            self.log_message.emit(
                f"✅ پردازش تمام شد! {found} شناسه با اطمینان بالا پیدا شد "
                f"(از {len(results)} ردیف)"
            )
            self.result_ready.emit(results)

        except Exception as e:
            self.log_message.emit(f"❌ خطای کلی: {str(e)}")
            if results:
                self.result_ready.emit(results)
            else:
                self.error_occurred.emit(str(e))


# ============================================================================
# پنجره اصلی
# ============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🔍 Smart ID Finder v6.1 — جستجوی تقویت‌شده")
        self.setMinimumSize(1550, 920)
        self.engine = SearchEngine()
        self.matcher = Matcher()
        self.headers = []
        self.rows = []
        self.kc = {}
        self.results = []
        self.manual_map = {}
        self.file_path = ""
        self.thread = None
        self.init_ui()

    def init_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        layout = QVBoxLayout(cw)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        self.setup_main_tab()
        self.setup_log_tab()
        self.sb = QStatusBar()
        self.setStatusBar(self.sb)
        self.sb.showMessage("آماده — نسخه ۶ (دقت بالا)")

    def setup_main_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # فایل
        fg = QGroupBox("📁 انتخاب فایل اکسل")
        fl = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("مسیر فایل اکسل...")
        browse = QPushButton("📂 انتخاب فایل")
        browse.clicked.connect(self.browse)
        browse.setStyleSheet(
            "QPushButton{background:#4CAF50;color:white;padding:10px 20px;"
            "border-radius:5px;font-weight:bold}"
        )
        fl.addWidget(self.file_edit)
        fl.addWidget(browse)
        fg.setLayout(fl)
        layout.addWidget(fg)

        # تنظیمات
        sg = QGroupBox("⚙️ تنظیمات دقت")
        sl = QHBoxLayout()
        sl.addWidget(QLabel("حداقل برای بررسی:"))
        self.th_spin = QSpinBox()
        self.th_spin.setRange(40, 95)
        self.th_spin.setValue(CONFIG['default_threshold'])
        self.th_spin.setSuffix("%")
        sl.addWidget(self.th_spin)

        sl.addWidget(QLabel("  قبول خودکار از:"))
        self.auto_spin = QSpinBox()
        self.auto_spin.setRange(60, 99)
        self.auto_spin.setValue(CONFIG['auto_accept_threshold'])
        self.auto_spin.setSuffix("%")
        sl.addWidget(self.auto_spin)

        self.start_btn = QPushButton("🚀 شروع جستجو")
        self.start_btn.clicked.connect(self.start)
        self.start_btn.setStyleSheet(
            "QPushButton{background:#2196F3;color:white;padding:12px 28px;"
            "border-radius:5px;font-weight:bold;font-size:14px}"
        )
        sl.addWidget(self.start_btn)

        self.pause_btn = QPushButton("⏸️ توقف موقت")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.setEnabled(False)
        self.pause_btn.setStyleSheet(
            "QPushButton{background:#FF9800;color:white;padding:12px 18px;"
            "border-radius:5px;font-weight:bold}"
        )
        sl.addWidget(self.pause_btn)

        self.stop_btn = QPushButton("⏹️ توقف کامل")
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(
            "QPushButton{background:#f44336;color:white;padding:12px 18px;"
            "border-radius:5px;font-weight:bold}"
        )
        sl.addWidget(self.stop_btn)

        sg.setLayout(sl)
        layout.addWidget(sg)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # جدول
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "بارکد", "نام کالا", "شناسه", "نام تطبیق یافته",
            "درصد", "وضعیت", "کلمات دستی", "ویرایش", "انتخاب کاندید", "عملیات"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.ctx_menu)
        layout.addWidget(self.table)

        # دکمه‌های پایین
        bl = QHBoxLayout()
        self.export_btn = QPushButton("📊 Export Excel")
        self.export_btn.clicked.connect(self.export)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet(
            "QPushButton{background:#9C27B0;color:white;padding:10px 20px;"
            "border-radius:5px;font-weight:bold}"
        )
        bl.addWidget(self.export_btn)

        self.accept_all_review_btn = QPushButton("✅ قبول همه موارد «نیاز به بررسی»")
        self.accept_all_review_btn.clicked.connect(self.accept_all_reviews)
        self.accept_all_review_btn.setEnabled(False)
        self.accept_all_review_btn.setStyleSheet(
            "QPushButton{background:#607D8B;color:white;padding:10px 16px;"
            "border-radius:5px;font-weight:bold}"
        )
        bl.addWidget(self.accept_all_review_btn)

        self.stats_lbl = QLabel("")
        self.stats_lbl.setStyleSheet("font-weight:bold;font-size:13px")
        bl.addWidget(self.stats_lbl)
        layout.addLayout(bl)

        self.tabs.addTab(tab, "🏠 اصلی")

    def setup_log_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 10))
        layout.addWidget(self.log_text)
        cb = QPushButton("🗑️ پاک کردن لاگ")
        cb.clicked.connect(lambda: self.log_text.clear())
        layout.addWidget(cb)
        self.tabs.addTab(tab, "📋 لاگ")

    def ctx_menu(self, pos):
        item = self.table.itemAt(pos)
        if not item:
            return
        row = item.row()
        menu = QMenu()
        a1 = QAction("✏️ ویرایش کلمات جستجو", self)
        a1.triggered.connect(lambda: self.edit_keywords(row))
        menu.addAction(a1)
        a2 = QAction("🔍 جستجوی مجدد", self)
        a2.triggered.connect(lambda: self.re_search(row))
        menu.addAction(a2)
        a3 = QAction("📋 انتخاب از کاندیدها", self)
        a3.triggered.connect(lambda: self.show_candidates(row))
        menu.addAction(a3)
        menu.exec_(self.table.mapToGlobal(pos))

    def edit_keywords(self, row):
        di = self.kc.get('description', -1)
        if di >= 0 and row < len(self.rows):
            pn = self.rows[row][di]
            cur = self.manual_map.get(row, "")
            dlg = ManualEditDialog(pn, cur, self)
            if dlg.exec_() == QDialog.Accepted:
                kw = dlg.get_keywords()
                if kw:
                    self.manual_map[row] = kw
                    self.table.setItem(row, 6, QTableWidgetItem(kw))
                    self.log(f"✏️ ردیف {row + 1}: کلمات دستی = '{kw}'")
                elif row in self.manual_map:
                    del self.manual_map[row]
                    self.table.setItem(row, 6, QTableWidgetItem(""))

    def show_candidates(self, row):
        """نمایش دیالوگ انتخاب از بین کاندیدهای ذخیره‌شده"""
        if row >= len(self.results):
            return
        r = self.results[row]
        cands = r.get('candidates') or []
        if not cands:
            QMessageBox.information(self, "اطلاع", "کاندیدی برای این ردیف ذخیره نشده. ابتدا جستجو کنید.")
            return

        di = self.kc.get('description', -1)
        pn = self.rows[row][di] if di >= 0 else ""

        # تبدیل به فرمت مورد نیاز دیالوگ
        top_list = []
        for c in cands:
            fake_score = {
                'final': c.get('score', 0),
                'details': c.get('details', '')
            }
            top_list.append((
                {'id': c['id'], 'name': c['name']},
                fake_score
            ))

        dlg = CandidateSelectDialog(pn, top_list, self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected:
            cid, cname, sim, details = dlg.selected
            self.table.setItem(row, 2, QTableWidgetItem(cid))
            self.table.setItem(row, 3, QTableWidgetItem(cname))
            self.table.setItem(row, 4, QTableWidgetItem(f"{sim:.1f}%"))
            st = "✅ انتخاب دستی"
            si = QTableWidgetItem(st)
            si.setBackground(QColor(180, 230, 255))
            self.table.setItem(row, 5, si)

            self.results[row] = {
                **r,
                'found_id': cid,
                'found_name': cname,
                'similarity': sim,
                'status': st,
                'score_details': details or r.get('score_details', '')
            }
            self.log(f"✅ ردیف {row + 1}: شناسه دستی انتخاب شد → {cid}")
            self._update_stats()

    def re_search(self, row):
        di = self.kc.get('description', -1)
        if di < 0 or row >= len(self.rows):
            return

        pn = self.rows[row][di]
        mkw = self.manual_map.get(row, "")

        self.log(f"\n🔍 جستجوی مجدد ردیف {row + 1}: {pn}")
        if mkw:
            self.log(f"  📝 دستی: {mkw}")

        if not self.engine.init_session(self.log):
            self.log("❌ خطا در اتصال")
            return

        cands, slog = self.engine.search(pn, mkw, self.log)
        if not cands:
            self.log("  ❌ نتیجه‌ای یافت نشد")
            QMessageBox.information(self, "نتیجه", "هیچ نتیجه‌ای یافت نشد!")
            return

        self.log(f"  📊 {len(cands)} نتیجه")
        best, scores, decision, top = self.matcher.decide(pn, cands)
        sim = scores.get('final', 0)

        cand_info = [
            {'id': c['id'], 'name': c['name'], 'score': s['final'],
             'details': s.get('details', '')}
            for c, s in top
        ]

        if decision == 'auto':
            st = "✅ تطابق عالی"
            fid = best['id']
        elif decision == 'review':
            st = "🔍 نیاز به بررسی دستی"
            fid = ''
        else:
            st = "❌ تطابق ضعیف / رد شده"
            fid = ''

        self.table.setItem(row, 2, QTableWidgetItem(fid))
        self.table.setItem(row, 3, QTableWidgetItem(best['name'] if best else ''))
        self.table.setItem(row, 4, QTableWidgetItem(f"{sim:.1f}%"))
        si = QTableWidgetItem(st)
        if sim >= 82:
            si.setBackground(QColor(200, 255, 200))
        elif sim >= 65:
            si.setBackground(QColor(255, 255, 200))
        else:
            si.setBackground(QColor(255, 200, 200))
        self.table.setItem(row, 5, si)

        if row < len(self.results):
            self.results[row] = {
                'found_id': fid,
                'found_name': best['name'] if best else '',
                'similarity': sim,
                'status': st,
                'search_info': ' | '.join(slog),
                'score_details': scores.get('details', ''),
                'candidates': cand_info,
                '_best_id': best['id'] if best else '',
                '_best_name': best['name'] if best else ''
            }

        self.log(f"  {st}: {best['name'][:50] if best else '-'} ({sim:.1f}%)")
        # باز کردن دیالوگ انتخاب برای راحتی کاربر
        self.show_candidates(row)

    def browse(self):
        fp, _ = QFileDialog.getOpenFileName(
            self, "انتخاب فایل", "", "Excel (*.xlsx *.xls)"
        )
        if fp:
            self.file_path = fp
            self.file_edit.setText(fp)
            self.load(fp)

    def load(self, fp):
        try:
            self.rows, self.headers, self.kc = ExcelHandler.read_excel(fp)
            self.log(f"✅ {len(self.rows)} ردیف بارگذاری شد")
            QMessageBox.information(self, "موفقیت", f"{len(self.rows)} ردیف بارگذاری شد")
        except Exception as e:
            QMessageBox.critical(self, "خطا", str(e))

    def start(self):
        if not self.rows:
            QMessageBox.warning(self, "هشدار", "فایل انتخاب نشده!")
            return
        if 'description' not in self.kc:
            QMessageBox.warning(self, "هشدار", "ستون LongDescription پیدا نشد!")
            return

        threshold = self.th_spin.value()
        auto_accept = self.auto_spin.value()
        if auto_accept < threshold:
            QMessageBox.warning(
                self, "هشدار",
                "مقدار «قبول خودکار» باید بزرگ‌تر یا مساوی «حداقل بررسی» باشد."
            )
            return

        self.matcher.threshold = threshold
        self.matcher.auto_accept = auto_accept

        self.start_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setText("⏸️ توقف موقت")
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.table.setRowCount(0)
        self.results = []
        self.accept_all_review_btn.setEnabled(False)

        self.thread = ProcessingThread(
            self.rows, self.kc, threshold, auto_accept,
            self.engine, self.matcher, self.manual_map,
            interactive_review=False
        )
        self.thread.progress_updated.connect(self._prog)
        self.thread.result_ready.connect(self._done)
        self.thread.error_occurred.connect(self._err)
        self.thread.log_message.connect(self.log)
        self.thread.start()

    def toggle_pause(self):
        if not self.thread:
            return
        if self.thread.is_paused:
            self.thread.resume()
            self.pause_btn.setText("⏸️ توقف موقت")
            self.pause_btn.setStyleSheet(
                "QPushButton{background:#FF9800;color:white;padding:12px 18px;"
                "border-radius:5px;font-weight:bold}"
            )
            self.log("▶️ ادامه...")
        else:
            self.thread.pause()
            self.pause_btn.setText("▶️ ادامه")
            self.pause_btn.setStyleSheet(
                "QPushButton{background:#4CAF50;color:white;padding:12px 18px;"
                "border-radius:5px;font-weight:bold}"
            )
            self.log("⏸️ توقف موقت")

    def stop_processing(self):
        if not self.thread:
            return
        self.thread.stop()
        self.log("⏹️ توقف کامل — نتایج تا اینجا لود می‌شوند...")

    def _prog(self, v, m):
        self.progress.setValue(v)
        self.sb.showMessage(m)

    def _done(self, results):
        self.results = results
        self.table.setRowCount(len(results))
        bi = self.kc.get('barcode', -1)
        di = self.kc.get('description', -1)
        found = 0
        review_count = 0

        for ri, r in enumerate(results):
            if ri < len(self.rows):
                row = self.rows[ri]
                bc = row[bi] if bi >= 0 and bi < len(row) else ""
                ds = row[di] if di >= 0 and di < len(row) else ""
                self.table.setItem(ri, 0, QTableWidgetItem(bc))
                self.table.setItem(ri, 1, QTableWidgetItem(ds))
                self.table.setItem(ri, 2, QTableWidgetItem(r.get('found_id', '')))
                self.table.setItem(ri, 3, QTableWidgetItem(r.get('found_name', '')))
                self.table.setItem(ri, 4, QTableWidgetItem(f"{r.get('similarity', 0):.1f}%"))

                sim = r.get('similarity', 0)
                st = r.get('status', '')
                si = QTableWidgetItem(st)
                if '✅' in st or sim >= 82:
                    si.setBackground(QColor(200, 255, 200))
                elif '🔍' in st or '🔶' in st or (65 <= sim < 82):
                    si.setBackground(QColor(255, 255, 200))
                    review_count += 1
                else:
                    si.setBackground(QColor(255, 200, 200))
                self.table.setItem(ri, 5, si)

                mkw = self.manual_map.get(ri, "")
                self.table.setItem(ri, 6, QTableWidgetItem(mkw))

                eb = QPushButton("✏️")
                eb.clicked.connect(lambda checked, row=ri: self.edit_keywords(row))
                self.table.setCellWidget(ri, 7, eb)

                cb = QPushButton("📋")
                cb.setToolTip("انتخاب از کاندیدها")
                cb.clicked.connect(lambda checked, row=ri: self.show_candidates(row))
                self.table.setCellWidget(ri, 8, cb)

                self.table.setItem(ri, 9, QTableWidgetItem(""))

                if r.get('found_id'):
                    found += 1

        self.start_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.export_btn.setEnabled(True)
        self.accept_all_review_btn.setEnabled(review_count > 0)
        self.progress.setVisible(False)
        self._update_stats()
        self.log(
            f"\n✅ خلاصه: {found} شناسه قطعی | "
            f"{review_count} مورد نیاز به بررسی دستی"
        )

    def _update_stats(self):
        found = sum(1 for r in self.results if r.get('found_id'))
        total = len(self.results) or 1
        review = sum(1 for r in self.results if '🔍' in r.get('status', ''))
        self.stats_lbl.setText(
            f"✅ {found} قطعی  |  🔍 {review} بررسی  |  "
            f"کل {len(self.results)}  ({found / total * 100:.1f}%)"
        )

    def accept_all_reviews(self):
        """قبول بهترین کاندید برای همه مواردی که «نیاز به بررسی» دارند"""
        count = 0
        for ri, r in enumerate(self.results):
            if '🔍' in r.get('status', '') and r.get('_best_id'):
                cid = r['_best_id']
                cname = r.get('_best_name', r.get('found_name', ''))
                sim = r.get('similarity', 0)
                self.table.setItem(ri, 2, QTableWidgetItem(cid))
                self.table.setItem(ri, 3, QTableWidgetItem(cname))
                st = "✅ قبول‌شده از بررسی"
                si = QTableWidgetItem(st)
                si.setBackground(QColor(180, 230, 255))
                self.table.setItem(ri, 5, si)
                self.results[ri] = {
                    **r,
                    'found_id': cid,
                    'found_name': cname,
                    'status': st
                }
                count += 1
        self.log(f"✅ {count} مورد «نیاز به بررسی» قبول شد")
        self.accept_all_review_btn.setEnabled(False)
        self._update_stats()
        QMessageBox.information(self, "انجام شد", f"{count} شناسه از موارد بررسی قبول شد.")

    def _err(self, e):
        QMessageBox.critical(self, "خطا", str(e))
        self.start_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.progress.setVisible(False)

    def export(self):
        if not self.results:
            return
        try:
            # همیشه کنار فایل اصلی ذخیره می‌شود (با timestamp)
            out = ExcelHandler.save_excel(
                self.file_path, self.headers, self.rows, self.results, self.kc
            )
            QMessageBox.information(self, "موفقیت", f"ذخیره شد:\n{out}")
            self.log(f"📊 خروجی ذخیره شد: {out}")
        except Exception as e:
            QMessageBox.critical(self, "خطا", str(e))

    def log(self, m):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {m}")


def main():
    logging.basicConfig(level=logging.INFO, handlers=[logging.StreamHandler()])
    app = QApplication(sys.argv)
    app.setFont(QFont("Tahoma", 10))
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
