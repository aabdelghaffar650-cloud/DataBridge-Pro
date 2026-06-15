"""
DataBridge — Config Manager
===========================
مركزي لإدارة إعدادات البرنامج (API Key + KPI Targets).
يحل محل load_config2/save_config2 في tab9 و _load_cfg7 في tab8.

الاستخدام:
    from modules.config_manager import load_config, save_config, get_api_key, save_api_key
    from modules.config_manager import get_kpi_targets, save_kpi_targets
"""

import os
import json
import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)

# ── مسار حفظ الإعدادات ──────────────────────────────────────────
APPDATA_DIR: str = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "DataBridge"
)
CONFIG_FILE: str = os.path.join(APPDATA_DIR, "config.json")

# ── القيم الافتراضية ─────────────────────────────────────────────
_DEFAULT_KPI_TARGETS: Dict[str, Dict[str, int]] = {
    k: {"annual": 0, "Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0}
    for k in ["hiv_tests", "positive", "referrals", "followups", "refusal"]
}

_DEFAULT_CONFIG: Dict[str, Any] = {
    "gemini_api_key": "",
    "kpi_targets": _DEFAULT_KPI_TARGETS,
}


# ── الدوال الأساسية ──────────────────────────────────────────────

def load_config() -> Dict[str, Any]:
    """
    تحميل الإعدادات من config.json.
    إذا لم يوجد الملف أو كان تالفاً، ترجع القيم الافتراضية.
    """
    os.makedirs(APPDATA_DIR, exist_ok=True)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                stored = json.load(f)
            # دمج مع الافتراضي لضمان وجود كل المفاتيح عند إضافة مفاتيح جديدة مستقبلاً
            merged = _DEFAULT_CONFIG.copy()
            merged.update(stored)
            # ضمان وجود كل مؤشرات KPI حتى لو كانت config قديمة
            merged_targets = _DEFAULT_KPI_TARGETS.copy()
            for k, v in stored.get("kpi_targets", {}).items():
                if k in merged_targets:
                    merged_targets[k].update(v)
            merged["kpi_targets"] = merged_targets
            return merged
        except Exception as e:
            logger.warning(f"[config_manager] فشل تحميل config.json: {e} — استخدام الافتراضي")
    return _DEFAULT_CONFIG.copy()


def save_config(cfg: Dict[str, Any]) -> None:
    """حفظ الإعدادات كاملةً في config.json."""
    os.makedirs(APPDATA_DIR, exist_ok=True)
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"[config_manager] فشل حفظ config.json: {e}")
        raise


# ── واجهة مختصرة لـ API Key ──────────────────────────────────────

def get_api_key() -> str:
    """إرجاع Gemini API Key المحفوظ (أو سلسلة فارغة)."""
    return load_config().get("gemini_api_key", "")


def save_api_key(key: str) -> None:
    """حفظ Gemini API Key مع الاحتفاظ بباقي الإعدادات."""
    cfg = load_config()
    cfg["gemini_api_key"] = key.strip()
    save_config(cfg)


def delete_api_key() -> None:
    """حذف Gemini API Key."""
    save_api_key("")


# ── واجهة مختصرة لـ KPI Targets ──────────────────────────────────

def get_kpi_targets() -> Dict[str, Dict[str, int]]:
    """
    إرجاع المستهدفات المحفوظة.
    الهيكل: { "hiv_tests": {"annual":0,"Q1":0,...}, ... }
    """
    return load_config().get("kpi_targets", _DEFAULT_KPI_TARGETS.copy())


def save_kpi_targets(targets: Dict[str, Dict[str, int]]) -> None:
    """حفظ المستهدفات مع الاحتفاظ بباقي الإعدادات."""
    cfg = load_config()
    cfg["kpi_targets"] = targets
    save_config(cfg)
