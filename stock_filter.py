"""
股票指标筛选工具 - 完整版（功能层 + 性能层 + 月KDJ）
启动命令：streamlit run stock_filter.py
"""
import streamlit as st
import akshare as ak
import pandas as pd
import numpy as np
from stockstats import wrap
import time
import os
import json
import sqlite3
import shutil
import io
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import warnings

warnings.filterwarnings("ignore")
st.set_page_config(page_title="股票指标筛选工具", page_icon="📈", layout="wide")

# ============================================================
# 全局配置
# ============================================================
DATA_DIR = "data/daily"
HISTORY_DB = "data/history.db"
WATCHLIST_FILE = "data/watchlist.json"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs("data", exist_ok=True)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def _safe_float(v):
    try:
        if v is None or v == "" or v == "--" or pd.isna(v):
            return None
        return float(v)
    except Exception:
        return None


# ============================================================
# SQLite 历史记录
# ============================================================
def init_history_db():
    conn = sqlite3.connect(HISTORY_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS screening_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_time TEXT,
            index_name TEXT,
            logic_mode TEXT,
            conditions TEXT,
            result_count INTEGER,
            results_json TEXT
        )
    """)
    conn.commit()
    conn.close()


def save_history(index_name, logic_mode, conditions, results):
    try:
        conn = sqlite3.connect(HISTORY_DB)
        conn.execute(
            "INSERT INTO screening_history (run_time, index_name, logic_mode, conditions, result_count, results_json) VALUES (?,?,?,?,?,?)",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), index_name, logic_mode,
             json.dumps(conditions, ensure_ascii=False, default=str),
             len(results), json.dumps(results, ensure_ascii=False, default=str)),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        st.warning(f"历史记录保存失败：{e}")


def load_history(limit=30):
    try:
        conn = sqlite3.connect(HISTORY_DB)
        df = pd.read_sql_query(
            f"SELECT id, run_time, index_name, logic_mode, conditions, result_count FROM screening_history ORDER BY id DESC LIMIT {limit}",
            conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


def _to_dict(v):
    """如果 v 是字符串形式的 dict，尝试解析成 dict；否则原样返回"""
    if isinstance(v, dict):
        return v
    if isinstance(v, str):
        try:
            import ast
            parsed = ast.literal_eval(v)
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            pass
    return {}


def summarize_conditions(conditions_json):
    """把保存的 conditions JSON 转成简短摘要文字（兼容旧数据的字符串嵌套）"""
    try:
        c = json.loads(conditions_json) if isinstance(conditions_json, str) else conditions_json
    except Exception:
        return ""
    if not isinstance(c, dict):
        return ""
    parts = []

    # 均线
    ma = _to_dict(c.get("ma_golden"))
    if ma.get("enabled"):
        parts.append(f"MA{ma.get('fast')}↑{ma.get('slow')}")
    if _to_dict(c.get("ma_bull")).get("enabled"):
        parts.append("MA多头")
    pma = _to_dict(c.get("price_vs_ma"))
    if pma.get("enabled"):
        parts.append(f"价格{'>' if pma.get('relation')=='above' else '<'}MA{pma.get('period')}")
    dev = _to_dict(c.get("ma_deviation"))
    if dev.get("enabled"):
        parts.append(f"偏离MA{dev.get('period')}{dev.get('op')}{dev.get('value')}%")

    # MACD
    mc = _to_dict(c.get("macd_cross"))
    if mc.get("enabled"):
        parts.append("MACD" + ("金叉" if mc.get("mode") == "golden" else "死叉"))
    mh = _to_dict(c.get("macd_hist"))
    if mh.get("enabled"):
        parts.append("MACD柱" + ("由负转正" if mh.get("mode") == "turn_pos" else "由正转负"))

    # BOLL
    bl = _to_dict(c.get("boll"))
    if bl.get("enabled"):
        parts.append(f"BOLL:{bl.get('mode')}")

    # KDJ 日/周/月（每个周期内部还有 K/D/J 三层嵌套）
    for period_key, label in [("kdj_daily", "日"), ("kdj_weekly", "周"), ("kdj_monthly", "月")]:
        kk = _to_dict(c.get(period_key))
        for k in ["K", "D", "J"]:
            sub = _to_dict(kk.get(k))
            if sub.get("enabled"):
                parts.append(f"{label}{k}{sub.get('op')}{sub.get('value')}")

    # KDJ金叉
    kc = _to_dict(c.get("kdj_cross"))
    if kc.get("enabled"):
        parts.append("KDJ" + ("金叉" if kc.get("mode") == "golden" else "死叉"))

    # RSI
    rsi = _to_dict(c.get("rsi"))
    if rsi.get("enabled"):
        parts.append(f"RSI{rsi.get('period')}{rsi.get('op')}{rsi.get('value')}")

    # 量
    vol = _to_dict(c.get("volume"))
    if vol.get("enabled"):
        parts.append(f"放量{vol.get('multiple')}倍")
    vp = _to_dict(c.get("vol_price"))
    if vp.get("enabled"):
        parts.append("量增价升" if vp.get("mode") == "up" else "量缩价跌")

    # ATR
    atr = _to_dict(c.get("atr"))
    if atr.get("enabled"):
        parts.append(f"ATR{atr.get('op')}{atr.get('value')}%")

    # 突破/新低
    bh = _to_dict(c.get("break_high"))
    if bh.get("enabled"):
        parts.append(f"突破{bh.get('n')}日新高")
    blw = _to_dict(c.get("break_low"))
    if blw.get("enabled"):
        parts.append(f"跌破{blw.get('n')}日新低")

    # 连续涨跌
    stk = _to_dict(c.get("streak"))
    if stk.get("enabled"):
        d = "涨" if stk.get("direction") == "up" else "跌"
        parts.append(f"连续{d}{stk.get('op')}{stk.get('days')}天")

    return " | ".join(parts) if parts else "（无条件）"


def load_history_detail(history_id):
    try:
        conn = sqlite3.connect(HISTORY_DB)
        df = pd.read_sql_query(
            "SELECT * FROM screening_history WHERE id=?", conn, params=(history_id,))
        conn.close()
        if df.empty:
            return None
        row = df.iloc[0]
        return {
            "run_time": row["run_time"],
            "index_name": row["index_name"],
            "conditions": json.loads(row["conditions"]) if row["conditions"] else {},
            "results": json.loads(row["results_json"]) if row["results_json"] else [],
        }
    except Exception:
        return None
def delete_history(history_id):
    """删除指定ID的历史记录"""
    try:
        conn = sqlite3.connect(HISTORY_DB)
        conn.execute("DELETE FROM screening_history WHERE id=?", (history_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        return False


def clear_all_history():
    """清空所有历史记录"""
    try:
        conn = sqlite3.connect(HISTORY_DB)
        conn.execute("DELETE FROM screening_history")
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False


# ============================================================
# 自选股管理
# ============================================================
def load_watchlist():
    if os.path.exists(WATCHLIST_FILE):
        try:
            with open(WATCHLIST_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_watchlist(lst):
    with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
        json.dump(lst, f, ensure_ascii=False)


# ============================================================
# 预设策略模板
# ============================================================
PRESET_STRATEGIES = {
    "超跌反弹": {
        "ui_dkdj_on": True, "ui_dk_k_on": True, "ui_dk_k_v": 20.0, "ui_dk_k_op": "<",
        "ui_dk_d_on": True, "ui_dk_d_v": 20.0, "ui_dk_d_op": "<",
        "ui_wkdj_on": True, "ui_wk_k_on": True, "ui_wk_k_v": 30.0, "ui_wk_k_op": "<",
        "ui_rsi_on": True, "ui_rsi_val": 30.0, "ui_rsi_op": "<",
        "rd_logic": "满足任意一个 (OR)",
    },
    "强势突破": {
        "ui_ma_bull": True,
        "ui_bh_on": True, "ui_bh_n": 60,
        "ui_vol_on": True, "ui_vol_mult": 1.5,
        "rd_logic": "同时满足全部 (AND)",
    },
    "低吸布局": {
        "fund_on": True,
        "pe_on": True, "pe_max": 20.0,
        "roe_on": True, "roe_min": 15.0,
        "rd_logic": "同时满足全部 (AND)",
    },
    "周月共振金叉": {
        "ui_wkdj_on": True, "ui_wk_k_on": True, "ui_wk_k_v": 30.0, "ui_wk_k_op": "<",
        "ui_mkdj_on": True, "ui_mk_k_on": True, "ui_mk_k_v": 40.0, "ui_mk_k_op": "<",
        "ui_kdj_cross_on": True, "ui_kdj_cross_mode": "金叉 (K上穿D)",
        "rd_logic": "同时满足全部 (AND)",
    },
}


# ============================================================
# 数据获取
# ============================================================
from concurrent.futures import ThreadPoolExecutor as _TPE
import time as _time_mod

def _fetch_with_timeout(func, timeout=8):
    """带超时执行函数，超时返回None，不阻塞主线程"""
    ex = _TPE(max_workers=1)
    fut = ex.submit(func)
    try:
        result = fut.result(timeout=timeout)
        ex.shutdown(wait=False)
        return result
    except Exception:
        ex.shutdown(wait=False)
        return None


@st.cache_data(ttl=300, show_spinner=False)
def get_index_components(index_code):
    """
    获取指数成分股：
    1. 本地缓存文件（30天内）→ 直接读
    2. 联网拉取（中证官网/新浪），每个源限时8秒
    3. 成功 → 保存到本地，下次直接读
    4. 全部失败但有旧本地文件 → 用旧文件兜底
    """
    cache_file = os.path.join("data", f"components_{index_code}.csv")

    # ---------- 优先本地缓存 ----------
    if os.path.exists(cache_file):
        try:
            age_days = (_time_mod.time() - os.path.getmtime(cache_file)) / 86400
            if age_days < 30:
                df = pd.read_csv(cache_file, dtype={"code": str})
                if not df.empty and "code" in df.columns and "name" in df.columns:
                    df["code"] = df["code"].astype(str).str.zfill(6)
                    return df[["code", "name"]].copy(), None
        except Exception:
            pass

    last_err = ""

    # ---------- 源1：中证指数官网（超时8秒） ----------
    def _try_csindex():
        return ak.index_stock_cons_csindex(symbol=index_code)
    try:
        df = _fetch_with_timeout(_try_csindex, timeout=8)
        if df is not None and not df.empty:
            df = df.rename(columns={"成分券代码": "code", "成分券名称": "name"})
            df["code"] = df["code"].astype(str).str.zfill(6)
            result = df[["code", "name"]].copy()
            try:
                result.to_csv(cache_file, index=False)
            except Exception:
                pass
            return result, None
        else:
            last_err = "中证官网超时"
    except Exception as e:
        last_err = f"中证官网: {str(e)[:60]}"

    # ---------- 源2：新浪（超时8秒） ----------
    def _try_sina():
        return ak.index_stock_cons(symbol=index_code)
    try:
        df = _fetch_with_timeout(_try_sina, timeout=8)
        if df is not None and not df.empty:
            rename_map = {}
            for c in df.columns:
                cs = str(c)
                if "代码" in cs: rename_map[c] = "code"
                if "名称" in cs: rename_map[c] = "name"
            df = df.rename(columns=rename_map)
            if "code" in df.columns and "name" in df.columns:
                df["code"] = df["code"].astype(str).str.zfill(6)
                result = df[["code", "name"]].copy()
                try:
                    result.to_csv(cache_file, index=False)
                except Exception:
                    pass
                return result, None
        else:
            last_err += " | 新浪超时"
    except Exception as e:
        last_err += f" | 新浪: {str(e)[:60]}"

    # ---------- 兜底：本地有旧文件就用 ----------
    if os.path.exists(cache_file):
        try:
            df = pd.read_csv(cache_file, dtype={"code": str})
            if not df.empty and "code" in df.columns and "name" in df.columns:
                df["code"] = df["code"].astype(str).str.zfill(6)
                return df[["code", "name"]].copy(), None
        except Exception:
            pass

    return None, f"全部源失败 - {last_err}"


def fetch_kline_from_network(stock_code, days=200):
    end_date = pd.Timestamp.today().strftime("%Y%m%d")
    start_date = (pd.Timestamp.today() - pd.Timedelta(days=days * 2)).strftime("%Y%m%d")
    last_err = ""
    for attempt in range(3):
        try:
            df = ak.stock_zh_a_hist(symbol=stock_code, period="daily",
                                    start_date=start_date, end_date=end_date, adjust="qfq")
            if df is not None and not df.empty:
                df = df.rename(columns={"日期": "date", "开盘": "open", "最高": "high",
                                        "最低": "low", "收盘": "close", "成交量": "volume"})
                df = df[["date", "open", "high", "low", "close", "volume"]].copy()
                df["date"] = pd.to_datetime(df["date"])
                return df.sort_values("date").reset_index(drop=True), None
        except Exception as e:
            last_err = f"东财: {str(e)[:50]}"
            time.sleep(2)
    try:
        prefix = "sh" if stock_code.startswith("6") else "sz"
        df = ak.stock_zh_a_daily(symbol=f"{prefix}{stock_code}", adjust="qfq")
        if df is not None and not df.empty:
            col_map = {}
            for c in df.columns:
                cl = str(c).lower().strip()
                if cl in ["date", "日期", "trade_date"]: col_map[c] = "date"
                elif cl in ["open", "开盘"]: col_map[c] = "open"
                elif cl in ["high", "最高"]: col_map[c] = "high"
                elif cl in ["low", "最低"]: col_map[c] = "low"
                elif cl in ["close", "收盘"]: col_map[c] = "close"
                elif cl in ["volume", "成交量", "vol"]: col_map[c] = "volume"
            df = df.rename(columns=col_map)
            if "date" not in df.columns:
                df = df.reset_index()
            needed = ["date", "open", "high", "low", "close", "volume"]
            missing = [n for n in needed if n not in df.columns]
            if missing:
                return None, f"新浪缺少列{missing}"
            df = df[needed].copy()
            df["date"] = pd.to_datetime(df["date"])
            df = df[df["date"] >= pd.to_datetime(start_date)]
            return df.sort_values("date").reset_index(drop=True), None
    except Exception as e:
        last_err += f" | 新浪: {str(e)[:50]}"
    return None, last_err
# ============================================================
# 【新增】交易日历 & 除权检测辅助函数
# ============================================================
@st.cache_data(ttl=86400, show_spinner=False)
def _get_trade_dates():
    """获取A股交易日历（用于判断缓存是否已是最新交易日）"""
    try:
        df = ak.tool_trade_date_hist_sina()
        return set(pd.to_datetime(df["trade_date"]).dt.normalize())
    except Exception:
        return None


def _check_seam_jump(old_df, new_df, tolerance=0.02):
    """
    检查新旧数据接缝处是否出现价格跳变（说明发生了除权）。
    old_df: 本地缓存    new_df: 新拉取的数据
    tolerance: 允许的最大相对差（默认2%）
    返回 True 表示检测到跳变，需要全量重拉
    """
    try:
        old_dates = set(old_df["date"])
        overlap = new_df[new_df["date"].isin(old_dates)]
        if overlap.empty or len(overlap) < 2:
            return False  # 没有重叠，无法判断
        merged = pd.merge(
            overlap[["date", "close"]],
            old_df[["date", "close"]],
            on="date", suffixes=("_new", "_old"),
        ).tail(3)
        if merged.empty:
            return False
        diffs = (merged["close_new"] - merged["close_old"]).abs() / merged["close_old"]
        return bool((diffs > tolerance).any())
    except Exception:
        return False

def load_or_update_kline(stock_code, days=200):
    """
    加载K线：本地缓存优先，只补拉缺失的近期数据。
    改进点：
      1. 用交易日历判断缓存是否已是最新交易日（不再用自然日）
      2. 检测新旧数据接缝处价格跳变（前复权除权）→ 触发全量重拉
      3. 缺口过大（>120天）→ 直接全量重拉，避免中间断层
      4. 原子写入，避免缓存损坏
    """
    file_path = os.path.join(DATA_DIR, f"{stock_code}.parquet")
    today = pd.Timestamp.today().normalize()

    def _tail(df):
        if len(df) > days:
            df = df.tail(days).reset_index(drop=True)
        return df

    def _save(df):
        try:
            tmp = file_path + ".tmp"
            df.to_parquet(tmp, index=False)
            os.replace(tmp, file_path)
        except Exception:
            pass

    # ---------- 本地已有缓存 ----------
    if os.path.exists(file_path):
        try:
            local_df = pd.read_parquet(file_path)
            local_df["date"] = pd.to_datetime(local_df["date"])
            local_df = local_df.sort_values("date").reset_index(drop=True)

            if not local_df.empty:
                last_cached = local_df["date"].max().normalize()

                # 用交易日历判断是否需要更新
                trade_dates = _get_trade_dates()
                if trade_dates is not None:
                    past_td = [d for d in trade_dates if d <= today]
                    latest_td = max(past_td) if past_td else today
                    need_update = last_cached < latest_td
                else:
                    # 兜底：拿不到日历时，自然日超过 3 天才更新
                    need_update = (today - last_cached).days > 3

                if not need_update:
                    return _tail(local_df), None

                gap_days = (today - last_cached).days

                # 缺口过大 → 全量重拉（避免中间断层）
                if gap_days > 120:
                    df, err = fetch_kline_from_network(stock_code, days=days)
                    if df is not None and not df.empty:
                        df = _tail(df)
                        _save(df)
                        return df, None
                    return _tail(local_df), None

                # 缺口较小 → 只补最近一段
                fetch_days = max(60, gap_days + 30)
                new_df, err = fetch_kline_from_network(stock_code, days=fetch_days)

                if new_df is not None and not new_df.empty:
                    # 检测除权跳变（前复权历史值被重算）
                    if _check_seam_jump(local_df, new_df):
                        df, err2 = fetch_kline_from_network(stock_code, days=days)
                        if df is not None and not df.empty:
                            df = _tail(df)
                            _save(df)
                            return df, None
                        return _tail(local_df), None

                    # 合并新旧数据
                    combined = pd.concat([local_df, new_df], ignore_index=True)
                    combined = combined.drop_duplicates(subset=["date"], keep="last")
                    combined = combined.sort_values("date").reset_index(drop=True)
                    combined = _tail(combined)
                    _save(combined)
                    return combined, None
                else:
                    # 补拉失败，用旧缓存兜底
                    return _tail(local_df), None
        except Exception:
            pass  # 缓存读失败，走下面全量拉取

    # ---------- 本地无缓存，全量拉取 ----------
    df, err = fetch_kline_from_network(stock_code, days=days)
    if df is not None and not df.empty:
        df = _tail(df)
        _save(df)
    return df, err


# ============================================================
# 基本面
# ============================================================
@st.cache_data(ttl=86400, show_spinner=False)
def get_fundamental_data(stock_code):
    result = {"PE": None, "PB": None, "总市值(亿)": None,
              "ROE(%)": None, "净利润增速(%)": None, "营收增速(%)": None}
    try:
        df = ak.stock_a_indicator_lg(symbol=stock_code)
        if df is not None and not df.empty:
            latest = df.iloc[-1]
            result["PE"] = _safe_float(latest.get("pe_ttm"))
            result["PB"] = _safe_float(latest.get("pb"))
            mv = _safe_float(latest.get("total_mv"))
            if mv is not None:
                result["总市值(亿)"] = round(mv / 1e8, 2)
    except Exception:
        pass
    try:
        df = ak.stock_financial_analysis_indicator(symbol=stock_code, start_year="2022")
        if df is not None and not df.empty:
            latest = df.iloc[0]
            for col in df.columns:
                cs = str(col)
                if "净资产收益率" in cs and "摊薄" not in cs and result["ROE(%)"] is None:
                    result["ROE(%)"] = _safe_float(latest[col])
                if "净利润增长率" in cs and result["净利润增速(%)"] is None:
                    result["净利润增速(%)"] = _safe_float(latest[col])
                if ("主营业务收入增长率" in cs or "营业收入增长率" in cs) and result["营收增速(%)"] is None:
                    result["营收增速(%)"] = _safe_float(latest[col])
    except Exception:
        pass
    return result


# ============================================================
# 辅助
# ============================================================
def is_special_stock(name):
    n = str(name).upper()
    if "ST" in n or "退" in n: return True
    if n.startswith("N"): return True
    return False


@st.cache_data(ttl=1800, show_spinner=False)
def get_suspended_stocks():
    try:
        df = ak.stock_zh_a_spot_em()
        if "交易状态" in df.columns:
            return set(df[df["交易状态"] != "交易"]["代码"].astype(str).str.zfill(6).tolist())
        if "最新价" in df.columns:
            return set(df[(df["最新价"].isna()) | (df["最新价"] == 0)]["代码"].astype(str).str.zfill(6).tolist())
        return set()
    except Exception:
        return set()


def to_weekly(kline_df):
    df = kline_df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df = df.set_index("date")
    weekly = df.resample("W-FRI").agg({"open": "first", "high": "max", "low": "min",
                                        "close": "last", "volume": "sum"}).dropna()
    return weekly.reset_index()

# ============================================================
# 【新增】通达信标准 KDJ 计算（与同花顺、通达信一致）
# ============================================================
def calc_kdj(df, n=9, m1=3, m2=3):
    """
    RSV = (CLOSE - LLV(LOW, N)) / (HHV(HIGH, N) - LLV(LOW, N)) * 100
    K   = SMA(RSV, M1, 1)
    D   = SMA(K,   M2, 1)
    J   = 3*K - 2*D
    说明：通达信 SMA(X, N, 1) 等价于 pandas 的 ewm(alpha=1/N, adjust=False)
    """
    out = df.copy()
    low_n  = out["low"].rolling(window=n, min_periods=1).min()
    high_n = out["high"].rolling(window=n, min_periods=1).max()

    denom = high_n - low_n
    rsv = (out["close"] - low_n) / denom.replace(0, np.nan) * 100
    rsv = rsv.fillna(50)  # 一字板/停牌等分母为0时，用中性值50

    k = rsv.ewm(alpha=1 / m1, adjust=False).mean()
    d = k.ewm(alpha=1 / m2, adjust=False).mean()
    j = 3 * k - 2 * d

    out["kdjk"] = k
    out["kdjd"] = d
    out["kdjj"] = j
    # 手工算金叉/死叉（stockstats 的 kdjk_xu_kdjd 我们不打算再用了）
    out["kdjk_xu_kdjd"] = (out["kdjk"] > out["kdjd"]) & (out["kdjk"].shift(1) <= out["kdjd"].shift(1))
    out["kdjd_xu_kdjk"] = (out["kdjd"] > out["kdjk"]) & (out["kdjd"].shift(1) <= out["kdjk"].shift(1))
    return out
def to_monthly(kline_df):
    """日K聚合为月K（月末）"""
    df = kline_df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df = df.set_index("date")
    # pandas 2.0+ 用 "ME"，旧版本用 "M"，这里做兼容
    try:
        monthly = df.resample("ME").agg({"open": "first", "high": "max", "low": "min",
                                          "close": "last", "volume": "sum"}).dropna()
    except Exception:
        monthly = df.resample("M").agg({"open": "first", "high": "max", "low": "min",
                                         "close": "last", "volume": "sum"}).dropna()
    return monthly.reset_index()


# ============================================================
# 评分
# ============================================================
def compute_score(hits, fund):
    score = 0
    details = []
    hit_str = " | ".join(hits)
    if "MA金叉" in hit_str: score += 30; details.append("MA金叉+30")
    if "MA多头" in hit_str: score += 25; details.append("MA多头+25")
    if "MACD金叉" in hit_str: score += 25; details.append("MACD金叉+25")
    if "MACD柱由负转正" in hit_str: score += 15; details.append("MACD柱反转+15")
    if "KDJ金叉" in hit_str: score += 25; details.append("KDJ金叉+25")
    if "周K" in hit_str: score += 20; details.append("周KDJ低位+20")
    if "月K" in hit_str: score += 25; details.append("月KDJ低位+25")
    if "RSI" in hit_str: score += 20; details.append("RSI条件+20")
    if "放量" in hit_str: score += 15; details.append("放量+15")
    if "突破" in hit_str and "新高" in hit_str: score += 20; details.append("突破新高+20")
    if "量增价升" in hit_str: score += 10; details.append("量增价升+10")
    if "跌破" in hit_str and "新低" in hit_str: score += 10; details.append("超跌+10")
    pe = fund.get("PE")
    roe = fund.get("ROE(%)")
    if pe is not None and 0 < pe < 30: score += 10; details.append("低PE+10")
    if roe is not None and roe > 15: score += 15; details.append("高ROE+15")
    return score, " ".join(details)


# ============================================================
# 核心判断
# ============================================================
def check_stock(kline_df, conds):
    if kline_df is None or len(kline_df) < 60:
        return [], 0, {}
    try:
        df = wrap(kline_df.copy())
        df.get("boll"); df.get("macd"); df.get("atr")
        # 用通达信标准算法覆盖 stockstats 的 KDJ
        kdj_df = calc_kdj(kline_df)
        df["kdjk"] = kdj_df["kdjk"].values
        df["kdjd"] = kdj_df["kdjd"].values
        df["kdjj"] = kdj_df["kdjj"].values
        df["kdjk_xu_kdjd"] = kdj_df["kdjk_xu_kdjd"].values
        df["kdjd_xu_kdjk"] = kdj_df["kdjd_xu_kdjk"].values
    except Exception:
        return [], 0, {}

    hits = []
    enabled_count = 0
    snapshot = {
        "最新价": round(df["close"].iloc[-1], 2),
        "涨跌幅(%)": round((df["close"].iloc[-1] / df["close"].iloc[-2] - 1) * 100, 2),
    }

    # 1. 均线金叉
    c = conds.get("ma_golden", {})
    if c.get("enabled"):
        enabled_count += 1
        fast, slow = int(c["fast"]), int(c["slow"])
        try:
            col_name = f"close_{fast}_sma_xu_close_{slow}_sma"
            _ = df[col_name]  # 显式赋值给 _ ，明确表示"只用一次，不显示"
            if df[col_name].iloc[-1]:
                hits.append(f"MA金叉({fast}↑{slow})")
        except Exception:
            pass

    # 2. 均线多头
    c = conds.get("ma_bull", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            m5, m10, m20 = df["close_5_sma"].iloc[-1], df["close_10_sma"].iloc[-1], df["close_20_sma"].iloc[-1]
            if m5 > m10 > m20:
                hits.append("MA多头(5>10>20)")
            snapshot["MA20"] = round(m20, 2)
        except Exception:
            pass

    # 3. 价格vs均线
    c = conds.get("price_vs_ma", {})
    if c.get("enabled"):
        enabled_count += 1
        period = int(c["period"])
        try:
            ma_col = f"close_{period}_sma"; _ = df[ma_col]
            close, ma_v = df["close"].iloc[-1], df[ma_col].iloc[-1]
            ok = (close > ma_v) if c["relation"] == "above" else (close < ma_v)
            if ok:
                hits.append(f"价格{'>' if c['relation']=='above' else '<'}MA{period}")
        except Exception:
            pass

    # 4. 偏离均线
    c = conds.get("ma_deviation", {})
    if c.get("enabled"):
        enabled_count += 1
        period = int(c["period"]); th = float(c["value"])
        try:
            ma_col = f"close_{period}_sma"; _ = df[ma_col]
            close, ma_v = df["close"].iloc[-1], df[ma_col].iloc[-1]
            if ma_v > 0:
                dev = (close - ma_v) / ma_v * 100
                ok = (dev > th) if c["op"] == ">" else (dev < th)
                if ok:
                    hits.append(f"偏离MA{period}{c['op']}{th}%(={dev:.1f})")
        except Exception:
            pass

    # 5. MACD金叉/死叉
    c = conds.get("macd_cross", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            if c["mode"] == "golden" and df["macd_xu_macds"].iloc[-1]:
                hits.append("MACD金叉")
            elif c["mode"] == "dead" and df["macds_xu_macd"].iloc[-1]:
                hits.append("MACD死叉")
        except Exception:
            pass

    # 6. MACD柱反转
    c = conds.get("macd_hist", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            h_now, h_prev = df["macdh"].iloc[-1], df["macdh"].iloc[-2]
            if c["mode"] == "turn_pos" and h_now > 0 and h_prev <= 0:
                hits.append("MACD柱由负转正")
            elif c["mode"] == "turn_neg" and h_now < 0 and h_prev >= 0:
                hits.append("MACD柱由正转负")
        except Exception:
            pass

    # 7. BOLL
    c = conds.get("boll", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            close, ub, lb, mid = df["close"].iloc[-1], df["boll_ub"].iloc[-1], df["boll_lb"].iloc[-1], df["boll"].iloc[-1]
            mode = c["mode"]
            if mode == "break_upper" and close > ub: hits.append("突破BOLL上轨")
            elif mode == "break_lower" and close < lb: hits.append("跌破BOLL下轨")
            elif mode == "above_mid" and close > mid: hits.append("价格在BOLL中轨上方")
            elif mode == "below_mid" and close < mid: hits.append("价格在BOLL中轨下方")
        except Exception:
            pass

    # 8. 日线KDJ数值
    c = conds.get("kdj_daily", {})
    for key, col in [("K", "kdjk"), ("D", "kdjd"), ("J", "kdjj")]:
        sub = c.get(key, {})
        if sub.get("enabled"):
            enabled_count += 1
            try:
                v = df[col].iloc[-1]
                ok = (v > sub["value"]) if sub["op"] == ">" else (v < sub["value"])
                if ok:
                   hits.append(f"日{key}{sub['op']}{sub['value']}(={v:.1f})")
                snapshot[f"日{key}"] = round(v, 2)
            except Exception:
                pass

    # 9. KDJ金叉/死叉
    c = conds.get("kdj_cross", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            if c["mode"] == "golden" and df["kdjk_xu_kdjd"].iloc[-1]:
                hits.append("KDJ金叉")
            elif c["mode"] == "dead" and df["kdjd_xu_kdjk"].iloc[-1]:
                hits.append("KDJ死叉")
        except Exception:
            pass

    # 10. 周线KDJ
    c = conds.get("kdj_weekly", {})
    if any(c.get(k, {}).get("enabled") for k in ["K", "D", "J"]):
        try:
            weekly = to_weekly(kline_df)
            if len(weekly) >= 12:
                wdf = calc_kdj(weekly)
                for key, col in [("K", "kdjk"), ("D", "kdjd"), ("J", "kdjj")]:
                    sub = c.get(key, {})
                    if sub.get("enabled"):
                        enabled_count += 1
                        v = wdf[col].iloc[-1]
                        ok = (v > sub["value"]) if sub["op"] == ">" else (v < sub["value"])
                        if ok:
                            hits.append(f"周{key}{sub['op']}{sub['value']}(={v:.1f})")
                        snapshot[f"周{key}"] = round(v, 2)
        except Exception:
            pass

    # 11. 月线KDJ（新增）
    c = conds.get("kdj_monthly", {})
    if any(c.get(k, {}).get("enabled") for k in ["K", "D", "J"]):
        try:
            monthly = to_monthly(kline_df)
            if len(monthly) >= 6:
                mdf = calc_kdj(monthly)
                for key, col in [("K", "kdjk"), ("D", "kdjd"), ("J", "kdjj")]:
                    sub = c.get(key, {})
                    if sub.get("enabled"):
                        enabled_count += 1
                        v = mdf[col].iloc[-1]
                        ok = (v > sub["value"]) if sub["op"] == ">" else (v < sub["value"])
                        if ok:
                            hits.append(f"月{key}{sub['op']}{sub['value']}(={v:.1f})")
                        snapshot[f"月{key}"] = round(v, 2)
        except Exception:
            pass

    # 12. RSI
    c = conds.get("rsi", {})
    if c.get("enabled"):
        enabled_count += 1
        period = int(c["period"])
        try:
            col = f"rsi_{period}"; _ = df[col]
            v = df[col].iloc[-1]
            ok = (v > c["value"]) if c["op"] == ">" else (v < c["value"])
            if ok:
                hits.append(f"RSI{period}{c['op']}{c['value']}(={v:.1f})")
            snapshot[f"RSI{period}"] = round(v, 2)
        except Exception:
            pass

    # 13. 成交量放大
    c = conds.get("volume", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            vol_5 = df["volume"].rolling(5).mean().iloc[-1]
            vol_now = df["volume"].iloc[-1]
            if vol_5 > 0 and vol_now > vol_5 * c["multiple"]:
                hits.append(f"放量{c['multiple']}倍")
            snapshot["量比(5日)"] = round(vol_now / vol_5, 2) if vol_5 > 0 else 0
        except Exception:
            pass

    # 14. 量价配合
    c = conds.get("vol_price", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            vol_5 = df["volume"].rolling(5).mean().iloc[-1]
            vol_now = df["volume"].iloc[-1]
            c_now, c_prev = df["close"].iloc[-1], df["close"].iloc[-2]
            mode = c["mode"]
            if mode == "up" and vol_5 > 0 and vol_now > vol_5 * 1.2 and c_now > c_prev:
                hits.append("量增价升")
            elif mode == "down" and vol_5 > 0 and vol_now < vol_5 * 0.8 and c_now < c_prev:
                hits.append("量缩价跌")
        except Exception:
            pass

    # 15. 突破N日新高
    c = conds.get("break_high", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            n = int(c["n"])
            if len(df) > n:
                prev_high = df["high"].iloc[-(n + 1):-1].max()
                if df["close"].iloc[-1] > prev_high:
                    hits.append(f"突破{n}日新高")
        except Exception:
            pass

    # 16. 跌破N日新低
    c = conds.get("break_low", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            n = int(c["n"])
            if len(df) > n:
                prev_low = df["low"].iloc[-(n + 1):-1].min()
                if df["close"].iloc[-1] < prev_low:
                    hits.append(f"跌破{n}日新低")
        except Exception:
            pass

    # 17. ATR
    c = conds.get("atr", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            atr_v = df["atr"].iloc[-1]
            close_now = df["close"].iloc[-1]
            atr_pct = atr_v / close_now * 100 if close_now > 0 else 0
            threshold = float(c["value"]); op = c["op"]
            ok = (atr_pct > threshold) if op == ">" else (atr_pct < threshold)
            if ok:
                hits.append(f"ATR{op}{threshold}%(={atr_pct:.2f})")
            snapshot["ATR%"] = round(atr_pct, 2)
        except Exception:
            pass
    # 18. 连续涨跌天数
    c = conds.get("streak", {})
    if c.get("enabled"):
        enabled_count += 1
        try:
            direction = c["direction"]
            op = c["op"]
            threshold = int(c["days"])
            closes = df["close"].values
            streak = 0
            for i in range(len(closes) - 1, 0, -1):
                diff = closes[i] - closes[i - 1]
                if direction == "up" and diff > 0:
                    streak += 1
                elif direction == "down" and diff < 0:
                    streak += 1
                else:
                    break
            ok = (streak >= threshold) if op == ">=" else (streak <= threshold)
            if ok:
                hits.append(f"连续{'涨' if direction == 'up' else '跌'}{streak}天")
            snapshot["连续天数"] = streak if direction == "up" else -streak
        except Exception:
            pass


    return hits, enabled_count, snapshot


# ============================================================
# 单股处理（供并发使用）
# ============================================================
def process_one_stock(code, name, history_days, conditions, use_and, skip_st, suspended_set):
    if skip_st:
        if is_special_stock(name) or code in suspended_set:
            return {"status": "skipped", "code": code, "name": name}
    kline, err = load_or_update_kline(code, days=history_days)
    if kline is None or len(kline) < 60:
        return {"status": "failed", "code": code, "name": name, "error": err or "数据不足60天"}
    hits, enabled_count, snap = check_stock(kline, conditions)
    if use_and:
        ok = (enabled_count > 0 and len(hits) >= enabled_count)
    else:
        ok = len(hits) > 0
    return {"status": "done", "code": code, "name": name, "hits": hits, "snap": snap, "ok": ok}


# ============================================================
# Excel 导出
# ============================================================
def export_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="筛选结果")
        ws = writer.sheets["筛选结果"]
        # 列宽
        for i, col in enumerate(df.columns, 1):
            letter = chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26)
            max_len = max(df[col].astype(str).map(len).max() if len(df) > 0 else 0, len(str(col))) + 2
            ws.column_dimensions[letter].width = min(max(max_len, 8), 35)
        # 涨跌幅颜色
        if "涨跌幅(%)" in df.columns:
            col_idx = list(df.columns).index("涨跌幅(%)") + 1
            for row in range(2, len(df) + 2):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.font = Font(color="C00000" if cell.value > 0 else "008000")
        # 表头
        for i in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=i)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
    output.seek(0)
    return output.getvalue()


# ============================================================
# 初始化
# ============================================================
init_history_db()

# ============================================================
# 页面
# ============================================================
st.title("📈 股票指标筛选工具")

INDEX_CODE_MAP = {"沪深300": "000300", "中证A50": "930050",
                  "中证A100": "000903", "中证A500": "000510"}

# 应用预设（必须在 widget 定义之前）
if "apply_preset" in st.session_state:
    preset = st.session_state.pop("apply_preset")
    for k, v in preset.items():
        st.session_state[k] = v

with st.sidebar:
    # ========================================================
    # 参数说明
    # ========================================================
    with st.expander("📖 参数说明（强烈建议先看一遍）", expanded=False):
        st.markdown("""
### 🎯 通用参数
- **股票池**：从哪个指数选股。沪深300大盘蓝筹、中证A50核心资产、中证A500中盘成长。**股票池越大，筛选时间越长**。
- **历史K线天数**：拉取多少天的日线。**越大越准，但越慢**。
  - 日线KDJ：≥80 天足够
  - 周线KDJ：≥150 天
  - **月线KDJ：≥250 天**（推荐值）
- **条件逻辑**：
  - **AND（同时满足）**：条件都要满足，结果少而精，适合严格筛选。
  - **OR（满足任意一个）**：任一条件满足即可，结果多，适合先广撒网。

### 📊 均线类
- **均线金叉**：短期均线上穿长期均线（如 MA5 上穿 MA20），典型买入信号。
  - 快线周期：短的那条，通常 5 或 10。
  - 慢线周期：长的那条，通常 20、30 或 60。
- **均线多头排列**：MA5 > MA10 > MA20，趋势向上的强势形态。
- **价格/均线关系**：价格在均线上方=强势，下方=弱势。
- **偏离均线幅度**：`(现价 - 均线)/均线 × 100%`。如"偏离20日均线>10%"表示短期涨得太急，有回调风险。

### 📉 MACD
- **金叉/死叉**：DIF 上穿 DEA 为金叉（看涨），下穿为死叉（看跌）。
- **柱状线反转**：MACD 柱由负转正（动能转多）或由正转负（动能转空），比金叉更早发生。

### 🎯 BOLL 布林带
由中轨（20日均线）+ 上下轨（±2倍标准差）组成：
- **突破上轨**：强势突破，可能加速
- **跌破下轨**：弱势破位，可能继续下跌
- **在中轨上方/下方**：多头/空头区域判断

### 🔱 KDJ（三种周期）
- **数值筛选**：K、D 的经典阈值
  - **K < 20**：超卖区（可能是买点）
  - **D < 20**：确认超卖
  - **J < 0**：极端超卖
  - **K > 80**：超买区（可能是卖点）
- **金叉/死叉**：K 上穿 D 是金叉（买入），K 下穿 D 是死叉（卖出）
- **日线/周线/月线**：
  - **日线**反映短期情绪，变化快
  - **周线**反映中期趋势，更可靠
  - **月线**反映长期方向，信号最强但触发很少
  - **共振**：如"周K<20 + 日线金叉"= 中期超卖+短期启动，是经典强信号

### 🌡️ RSI 相对强弱
范围 0~100：
- **RSI < 30**：超卖，可能反弹
- **RSI > 70**：超买，可能回调
- 常用周期：6（敏感）、12、14（稳健）

### 📦 成交量 / 形态 / ATR
- **放量倍数**：`今日量 / 5日均量`。1.5倍=温和放量，3倍以上=异动。
- **量增价升**：量能放大 + 价格上涨，健康上涨。
- **量缩价跌**：量能萎缩 + 价格下跌，缩量调整。
- **ATR%**：`ATR / 收盘价 × 100%`，代表日均波动幅度。3% 以上=活跃股，1% 以下=低波动股。
- **突破N日新高**：收盘价 > 过去N日最高。突破越长期的高点，信号越强。
- **跌破N日新低**：收盘价 < 过去N日最低。配合低位 KDJ 是超跌抄底信号。

### 💰 基本面
- **PE(TTM)**：市盈率，越低越便宜。0~30 是低估值区间，>100 需警惕。
- **PB**：市净率，< 3 比较健康。
- **ROE**：净资产收益率，> 15% 说明公司赚钱能力强。
- **总市值**：过滤掉太小的公司（如 < 100亿），避免流动性和退市风险。
- **净利润/营收增速**：> 0 表示同比增长，> 20% 表示高成长。

### ⚡ 性能设置
- **并发拉取**：多线程同时拉数据，快 3~5 倍，但**容易被限流**。默认关闭。
- **建议**：日常串行（稳），偶尔赶时间开 3 线程。

### 🏆 综合评分
每只命中的股票自动打分：
- 均线金叉 +30，MACD金叉 +25，KDJ金叉 +25
- 周线KDJ低位 +20，月线KDJ低位 +25
- 突破新高 +20，放量 +15
- 低PE(<30) +10，高ROE(>15) +15
- **分数越高，信号越强**。结果自动按分数从高到低排序。
        """)

    st.header("⚙️ 设置")
    index_name = st.selectbox("股票池", list(INDEX_CODE_MAP.keys()), index=0, key="sel_idx")
    history_days = st.slider("历史K线天数", 80, 600, 250, 10, key="sld_days")
    logic_mode = st.radio("条件逻辑", ["同时满足全部 (AND)", "满足任意一个 (OR)"], index=0, key="rd_logic")

    # ---------- 预设策略 ----------
    with st.expander("🎯 预设策略模板", expanded=False):
        st.caption("点击一键应用常用组合，再按需微调")
        for preset_name in PRESET_STRATEGIES.keys():
            if st.button(f"应用「{preset_name}」", key=f"btn_preset_{preset_name}"):
                st.session_state["apply_preset"] = PRESET_STRATEGIES[preset_name]
                st.rerun()

    # ---------- 自选股 ----------
    with st.expander("⭐ 自选股 / 关注列表", expanded=False):
        wl = load_watchlist()
        st.caption(f"当前自选 {len(wl)} 只")
        only_watchlist = st.checkbox("只在自选股中筛选", value=False, key="only_wl")
        new_code = st.text_input("添加自选（代码，逗号分隔）", key="wl_add_input")
        ca, cb = st.columns(2)
        with ca:
            if st.button("➕ 添加", key="btn_wl_add") and new_code:
                codes = [c.strip().zfill(6) for c in new_code.replace("，", ",").split(",") if c.strip()]
                wl_new = list(set(wl + codes))
                save_watchlist(wl_new)
                st.success(f"已添加，当前 {len(wl_new)} 只")
                st.rerun()
        with cb:
            if st.button("🗑️ 清空自选", key="btn_wl_clear"):
                save_watchlist([])
                st.success("已清空")
                st.rerun()
        if wl:
            with st.expander(f"查看自选列表（{len(wl)}）", expanded=False):
                st.write("、".join(wl))

    # ---------- 性能 ----------
    with st.expander("⚡ 性能设置", expanded=False):
        use_concurrent = st.checkbox("启用并发拉取（更快，但可能被限流）", value=False, key="perf_concurrent")
        num_workers = st.slider("并发线程数", 2, 5, 3, 1, key="perf_workers", disabled=not use_concurrent)

    # ---------- 数据管理 ----------
    with st.expander("🗄️ 数据管理", expanded=False):
        skip_st = st.checkbox("自动过滤 ST / 停牌 / 新股", value=True, key="chk_skip")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("📊 查看缓存", key="btn_view_cache"):
                if os.path.exists(DATA_DIR):
                    files = [f for f in os.listdir(DATA_DIR) if f.endswith(".parquet")]
                    st.session_state["cache_count"] = len(files)
        with col_b:
            if st.button("🗑️ 清空缓存", key="btn_clear_cache"):
                if os.path.exists(DATA_DIR):
                    shutil.rmtree(DATA_DIR); os.makedirs(DATA_DIR, exist_ok=True)
                    st.session_state["cache_msg"] = "缓存已清空"
        if "cache_count" in st.session_state:
            st.info(f"本地已缓存 {st.session_state['cache_count']} 只")
        if "cache_msg" in st.session_state:
            st.success(st.session_state["cache_msg"])

    # ---------- 基本面 ----------
    with st.expander("💰 基本面筛选", expanded=False):
        fund_on = st.checkbox("启用基本面筛选", value=False, key="fund_on")
        pe_on = st.checkbox("PE(TTM) 区间", value=False, key="pe_on", disabled=not fund_on)
        pc1, pc2 = st.columns(2)
        with pc1:
            pe_min = st.number_input("PE最小", 0.0, 200.0, 0.0, 1.0, key="pe_min", disabled=not (fund_on and pe_on))
        with pc2:
            pe_max = st.number_input("PE最大", 0.0, 200.0, 30.0, 1.0, key="pe_max", disabled=not (fund_on and pe_on))
        pb_on = st.checkbox("PB 区间", value=False, key="pb_on", disabled=not fund_on)
        bc1, bc2 = st.columns(2)
        with bc1:
            pb_min = st.number_input("PB最小", 0.0, 50.0, 0.0, 0.1, key="pb_min", disabled=not (fund_on and pb_on))
        with bc2:
            pb_max = st.number_input("PB最大", 0.0, 50.0, 5.0, 0.1, key="pb_max", disabled=not (fund_on and pb_on))
        roe_on = st.checkbox("ROE 下限", value=False, key="roe_on", disabled=not fund_on)
        roe_min = st.number_input("ROE ≥ (%)", 0.0, 100.0, 10.0, 1.0, key="roe_min", disabled=not (fund_on and roe_on))
        mv_on = st.checkbox("总市值下限", value=False, key="mv_on", disabled=not fund_on)
        mv_min = st.number_input("总市值 ≥ (亿)", 0.0, 100000.0, 100.0, 10.0, key="mv_min", disabled=not (fund_on and mv_on))
        profit_on = st.checkbox("净利润增速 下限", value=False, key="profit_on", disabled=not fund_on)
        profit_min = st.number_input("净利润增速 ≥ (%)", -200.0, 500.0, 0.0, 5.0, key="profit_min", disabled=not (fund_on and profit_on))
        rev_on = st.checkbox("营收增速 下限", value=False, key="rev_on", disabled=not fund_on)
        rev_min = st.number_input("营收增速 ≥ (%)", -100.0, 500.0, 0.0, 5.0, key="rev_min", disabled=not (fund_on and rev_on))

    # ---------- 均线 ----------
    with st.expander("📊 均线类", expanded=False):
        ma_golden_on = st.checkbox("启用均线金叉", value=False, key="ui_ma_g_on")
        c1, c2 = st.columns(2)
        with c1:
            ma_fast = st.number_input("快线周期", 3, 60, 5, 1, key="ui_ma_fast", disabled=not ma_golden_on)
        with c2:
            ma_slow = st.number_input("慢线周期", 5, 120, 20, 1, key="ui_ma_slow", disabled=not ma_golden_on)
        ma_bull_on = st.checkbox("均线多头排列 (5>10>20)", value=False, key="ui_ma_bull")
        pma_on = st.checkbox("价格 / 均线关系", value=False, key="ui_pma_on")
        pma_period = st.number_input("均线周期##pma", 5, 120, 60, 1, key="ui_pma_period", disabled=not pma_on)
        pma_rel = st.radio("价格位置", ["在均线上方", "在均线下方"], horizontal=True, key="ui_pma_rel", disabled=not pma_on)
        dev_on = st.checkbox("价格偏离均线幅度", value=False, key="ui_dev_on")
        dev_period = st.number_input("参考均线##dev", 5, 120, 20, 1, key="ui_dev_period", disabled=not dev_on)
        dc1, dc2 = st.columns(2)
        with dc1:
            dev_op = st.selectbox("偏离条件##dev", [">", "<"], key="ui_dev_op", disabled=not dev_on)
        with dc2:
            dev_val = st.number_input("偏离%##dev", -50.0, 50.0, 5.0, 0.5, key="ui_dev_val", disabled=not dev_on)

    # ---------- MACD ----------
    with st.expander("📉 MACD", expanded=False):
        macd_on = st.checkbox("MACD 金叉/死叉", value=False, key="ui_macd_on")
        macd_mode = st.radio("形态", ["金叉 (DIF上穿DEA)", "死叉 (DEA上穿DIF)"], horizontal=True, key="ui_macd_mode", disabled=not macd_on)
        macdh_on = st.checkbox("MACD 柱状线反转", value=False, key="ui_macdh_on")
        macdh_mode = st.radio("方向", ["由负转正", "由正转负"], horizontal=True, key="ui_macdh_mode", disabled=not macdh_on)

    # ---------- BOLL ----------
    with st.expander("🎯 布林带 BOLL", expanded=False):
        boll_on = st.checkbox("启用 BOLL", value=False, key="ui_boll_on")
        boll_mode = st.selectbox("BOLL形态", ["突破上轨", "跌破下轨", "在中轨上方", "在中轨下方"], key="ui_boll_mode", disabled=not boll_on)

    # ---------- 日线 KDJ ----------
    with st.expander("🔱 日线 KDJ", expanded=False):
        dk_on = st.checkbox("日线 KDJ 数值筛选", value=False, key="ui_dkdj_on")
        dk_k_on = st.checkbox("K 值", value=False, key="ui_dk_k_on", disabled=not dk_on)
        kc1, kc2 = st.columns(2)
        with kc1:
            dk_k_op = st.selectbox("K条件##dk", [">", "<"], key="ui_dk_k_op", disabled=not dk_k_on)
        with kc2:
            dk_k_v = st.number_input("K阈值##dk", 0.0, 200.0, 80.0, 1.0, key="ui_dk_k_v", disabled=not dk_k_on)
        dk_d_on = st.checkbox("D 值", value=False, key="ui_dk_d_on", disabled=not dk_on)
        dc1, dc2 = st.columns(2)
        with dc1:
            dk_d_op = st.selectbox("D条件##dk", [">", "<"], key="ui_dk_d_op", disabled=not dk_d_on)
        with dc2:
            dk_d_v = st.number_input("D阈值##dk", 0.0, 200.0, 20.0, 1.0, key="ui_dk_d_v", disabled=not dk_d_on)
        dk_j_on = st.checkbox("J 值", value=False, key="ui_dk_j_on", disabled=not dk_on)
        jc1, jc2 = st.columns(2)
        with jc1:
            dk_j_op = st.selectbox("J条件##dk", [">", "<"], key="ui_dk_j_op", disabled=not dk_j_on)
        with jc2:
            dk_j_v = st.number_input("J阈值##dk", -100.0, 200.0, 0.0, 1.0, key="ui_dk_j_v", disabled=not dk_j_on)
        kdj_cross_on = st.checkbox("KDJ 金叉/死叉", value=False, key="ui_kdj_cross_on")
        kdj_cross_mode = st.radio("形态##kdjc", ["金叉 (K上穿D)", "死叉 (D上穿K)"], horizontal=True, key="ui_kdj_cross_mode", disabled=not kdj_cross_on)

    # ---------- 周线 KDJ ----------
    with st.expander("🔱 周线 KDJ", expanded=False):
        wk_on = st.checkbox("周线 KDJ 数值筛选", value=False, key="ui_wkdj_on")
        wk_k_on = st.checkbox("周K值", value=False, key="ui_wk_k_on", disabled=not wk_on)
        wkc1, wkc2 = st.columns(2)
        with wkc1:
            wk_k_op = st.selectbox("周K条件", [">", "<"], key="ui_wk_k_op", disabled=not wk_k_on)
        with wkc2:
            wk_k_v = st.number_input("周K阈值", 0.0, 200.0, 80.0, 1.0, key="ui_wk_k_v", disabled=not wk_k_on)
        wk_d_on = st.checkbox("周D值", value=False, key="ui_wk_d_on", disabled=not wk_on)
        wdc1, wdc2 = st.columns(2)
        with wdc1:
            wk_d_op = st.selectbox("周D条件", [">", "<"], key="ui_wk_d_op", disabled=not wk_d_on)
        with wdc2:
            wk_d_v = st.number_input("周D阈值", 0.0, 200.0, 20.0, 1.0, key="ui_wk_d_v", disabled=not wk_d_on)
        wk_j_on = st.checkbox("周J值", value=False, key="ui_wk_j_on", disabled=not wk_on)
        wjc1, wjc2 = st.columns(2)
        with wjc1:
            wk_j_op = st.selectbox("周J条件", [">", "<"], key="ui_wk_j_op", disabled=not wk_j_on)
        with wjc2:
            wk_j_v = st.number_input("周J阈值", -100.0, 200.0, 0.0, 1.0, key="ui_wk_j_v", disabled=not wk_j_on)

    # ---------- 月线 KDJ（新增）----------
    with st.expander("🔱 月线 KDJ", expanded=False):
        st.caption("月K数据需≥6个月，建议历史K线天数≥250")
        mk_on = st.checkbox("月线 KDJ 数值筛选", value=False, key="ui_mkdj_on")
        mk_k_on = st.checkbox("月K值", value=False, key="ui_mk_k_on", disabled=not mk_on)
        mkc1, mkc2 = st.columns(2)
        with mkc1:
            mk_k_op = st.selectbox("月K条件", [">", "<"], key="ui_mk_k_op", disabled=not mk_k_on)
        with mkc2:
            mk_k_v = st.number_input("月K阈值", 0.0, 200.0, 40.0, 1.0, key="ui_mk_k_v", disabled=not mk_k_on)
        mk_d_on = st.checkbox("月D值", value=False, key="ui_mk_d_on", disabled=not mk_on)
        mdc1, mdc2 = st.columns(2)
        with mdc1:
            mk_d_op = st.selectbox("月D条件", [">", "<"], key="ui_mk_d_op", disabled=not mk_d_on)
        with mdc2:
            mk_d_v = st.number_input("月D阈值", 0.0, 200.0, 40.0, 1.0, key="ui_mk_d_v", disabled=not mk_d_on)
        mk_j_on = st.checkbox("月J值", value=False, key="ui_mk_j_on", disabled=not mk_on)
        mjc1, mjc2 = st.columns(2)
        with mjc1:
            mk_j_op = st.selectbox("月J条件", [">", "<"], key="ui_mk_j_op", disabled=not mk_j_on)
        with mjc2:
            mk_j_v = st.number_input("月J阈值", -100.0, 200.0, 20.0, 1.0, key="ui_mk_j_v", disabled=not mk_j_on)

    # ---------- RSI ----------
    with st.expander("🌡️ RSI", expanded=False):
        rsi_on = st.checkbox("启用 RSI", value=False, key="ui_rsi_on")
        rc1, rc2, rc3 = st.columns(3)
        with rc1:
            rsi_period = st.number_input("周期##rsi", 2, 30, 6, 1, key="ui_rsi_period", disabled=not rsi_on)
        with rc2:
            rsi_op = st.selectbox("条件##rsi", [">", "<"], key="ui_rsi_op", disabled=not rsi_on)
        with rc3:
            rsi_val = st.number_input("阈值##rsi", 0.0, 100.0, 30.0, 1.0, key="ui_rsi_val", disabled=not rsi_on)

    # ---------- 成交量 / 形态 / ATR ----------
    with st.expander("📦 成交量 / 形态 / ATR", expanded=False):
        vol_on = st.checkbox("成交量放大", value=False, key="ui_vol_on")
        vol_mult = st.number_input("放量倍数", 1.0, 10.0, 1.5, 0.1, key="ui_vol_mult", disabled=not vol_on)
        vp_on = st.checkbox("量价配合", value=False, key="ui_vp_on")
        vp_mode = st.radio("形态##vp", ["量增价升", "量缩价跌"], horizontal=True, key="ui_vp_mode", disabled=not vp_on)
        atr_on = st.checkbox("ATR 波动率", value=False, key="ui_atr_on")
        ac1, ac2 = st.columns(2)
        with ac1:
            atr_op = st.selectbox("条件##atr", [">", "<"], key="ui_atr_op", disabled=not atr_on)
        with ac2:
            atr_val = st.number_input("ATR%阈值##atr", 0.0, 30.0, 3.0, 0.5, key="ui_atr_val", disabled=not atr_on)
        bh_on = st.checkbox("突破 N 日新高", value=False, key="ui_bh_on")
        bh_n = st.number_input("N 日##bh", 5, 250, 60, 5, key="ui_bh_n", disabled=not bh_on)
        bl_on = st.checkbox("跌破 N 日新低", value=False, key="ui_bl_on")
        bl_n = st.number_input("N 日##bl", 5, 250, 60, 5, key="ui_bl_n", disabled=not bl_on)
        
        # 连续涨跌
        streak_on = st.checkbox("连续涨/跌天数", value=False, key="ui_streak_on")
        streak_direction = st.radio("方向##streak", ["连续上涨", "连续下跌"],
                                    horizontal=True, key="ui_streak_dir",
                                    disabled=not streak_on)
        sc1, sc2 = st.columns(2)
        with sc1:
            streak_op = st.selectbox("条件##streak", [">=", "<="],
                                     key="ui_streak_op", disabled=not streak_on)
        with sc2:
            streak_days = st.number_input("天数##streak", 1, 30, 3, 1,
                                          key="ui_streak_days", disabled=not streak_on)

    # ---------- 单只股票体检 ----------
    st.divider()
    st.subheader("🧪 单只股票体检")
    test_input = st.text_input("输入股票代码或名称", key="test_code")
    test_button = st.button("开始体检", key="test_btn")
    if test_button and test_input:
        test_input = test_input.strip()
        test_code = None
        if test_input.isdigit() and len(test_input) == 6:
            test_code = test_input
        else:
            with st.spinner(f"正在查找「{test_input}」..."):
                try:
                    all_stocks = ak.stock_info_a_code_name()
                    match = all_stocks[all_stocks["name"].str.contains(test_input, na=False)]
                    if not match.empty:
                        test_code = str(match.iloc[0]["code"]).zfill(6)
                        st.info(f"✅ {match.iloc[0]['name']} → {test_code}")
                    else:
                        st.error(f"❌ 找不到「{test_input}」")
                except Exception as e:
                    st.error(f"❌ {e}")
        if test_code:
            with st.spinner("拉取中..."):
                t_df, t_err = load_or_update_kline(test_code, days=history_days)
                t_fund = get_fundamental_data(test_code)
            if t_df is None or len(t_df) < 60:
                st.error(f"❌ 拉取失败：{t_err}")
            else:
                try:
                    tdf = wrap(t_df.copy())
                    tdf.get("macd"); tdf.get("atr"); _ = tdf["rsi_6"]
                    tdf_kdj = calc_kdj(t_df)
                    tdf["kdjk"] = tdf_kdj["kdjk"].values
                    tdf["kdjd"] = tdf_kdj["kdjd"].values
                    tdf["kdjj"] = tdf_kdj["kdjj"].values

                    tw  = to_weekly(t_df)
                    wdf = calc_kdj(tw)

                    tm  = to_monthly(t_df)
                    mdf = calc_kdj(tm)

                    st.success(f"✅ {t_df['date'].max().strftime('%Y-%m-%d')}")
                    chart_df = t_df.set_index("date")[["close"]].copy(); chart_df.columns = ["收盘价"]
                    st.line_chart(chart_df, height=160)
                    r1 = st.columns(5)
                    r1[0].metric("最新价", round(t_df["close"].iloc[-1], 2))
                    r1[1].metric("日K", round(tdf["kdjk"].iloc[-1], 2))
                    r1[2].metric("日D", round(tdf["kdjd"].iloc[-1], 2))
                    r1[3].metric("周K", round(wdf["kdjk"].iloc[-1], 2) if len(tw) >= 12 else "N/A")
                    r1[4].metric("周D", round(wdf["kdjd"].iloc[-1], 2) if len(tw) >= 12 else "N/A")
                    r2 = st.columns(5)
                    r2[0].metric("月K", round(mdf["kdjk"].iloc[-1], 2) if len(tm) >= 6 else "N/A")
                    r2[1].metric("月D", round(mdf["kdjd"].iloc[-1], 2) if len(tm) >= 6 else "N/A")
                    r2[2].metric("RSI6", round(tdf["rsi_6"].iloc[-1], 2))
                    r2[3].metric("ATR%", round(tdf["atr"].iloc[-1] / t_df["close"].iloc[-1] * 100, 2))
                    r2[4].metric("PE", t_fund["PE"] if t_fund["PE"] is not None else "N/A")
                except Exception as e:
                    st.error(f"❌ {e}")


# ============================================================
# 条件概览
# ============================================================
st.subheader("📋 当前启用的筛选条件")
active = []
if fund_on: active.append("💰 基本面筛选")
if ma_golden_on: active.append(f"✅ MA金叉 {ma_fast}↑{ma_slow}")
if ma_bull_on:   active.append("✅ MA多头")
if pma_on:       active.append(f"✅ 价格{pma_rel}MA{pma_period}")
if dev_on:       active.append(f"✅ 偏离MA{dev_period}{dev_op}{dev_val}%")
if macd_on:      active.append(f"✅ MACD{macd_mode}")
if macdh_on:     active.append(f"✅ MACD柱{macdh_mode}")
if boll_on:      active.append(f"✅ BOLL{boll_mode}")
if dk_on and dk_k_on: active.append(f"✅ 日K{dk_k_op}{dk_k_v}")
if dk_on and dk_d_on: active.append(f"✅ 日D{dk_d_op}{dk_d_v}")
if dk_on and dk_j_on: active.append(f"✅ 日J{dk_j_op}{dk_j_v}")
if kdj_cross_on: active.append(f"✅ KDJ{kdj_cross_mode}")
if wk_on and wk_k_on: active.append(f"✅ 周K{wk_k_op}{wk_k_v}")
if wk_on and wk_d_on: active.append(f"✅ 周D{wk_d_op}{wk_d_v}")
if wk_on and wk_j_on: active.append(f"✅ 周J{wk_j_op}{wk_j_v}")
if mk_on and mk_k_on: active.append(f"✅ 月K{mk_k_op}{mk_k_v}")
if mk_on and mk_d_on: active.append(f"✅ 月D{mk_d_op}{mk_d_v}")
if mk_on and mk_j_on: active.append(f"✅ 月J{mk_j_op}{mk_j_v}")
if rsi_on:       active.append(f"✅ RSI{rsi_period}{rsi_op}{rsi_val}")
if vol_on:       active.append(f"✅ 放量{vol_mult}倍")
if vp_on:        active.append(f"✅ {vp_mode}")
if atr_on:       active.append(f"✅ ATR{atr_op}{atr_val}%")
if bh_on:        active.append(f"✅ 突破{bh_n}日新高")
if bl_on:        active.append(f"✅ 跌破{bl_n}日新低")
if streak_on:    active.append(f"✅ {streak_direction} {streak_op} {streak_days}天")

if not active:
    st.warning("⚠️ 未启用任何筛选条件")
else:
    for a in active: st.write(a)
    st.caption(f"筛选逻辑：{logic_mode}")

st.divider()
start_button = st.button("🚀 开始筛选", type="primary", use_container_width=True)


# ============================================================
# 执行筛选
# ============================================================
if start_button:
    if not active:
        st.error("请先勾选至少一个筛选条件。")
        st.stop()

    conditions = {
        "ma_golden": {"enabled": ma_golden_on, "fast": ma_fast, "slow": ma_slow},
        "ma_bull":   {"enabled": ma_bull_on},
        "price_vs_ma": {"enabled": pma_on, "period": pma_period, "relation": "above" if pma_rel == "在均线上方" else "below"},
        "ma_deviation": {"enabled": dev_on, "period": dev_period, "op": dev_op, "value": dev_val},
        "macd_cross": {"enabled": macd_on, "mode": "golden" if "金叉" in macd_mode else "dead"},
        "macd_hist": {"enabled": macdh_on, "mode": "turn_pos" if macdh_mode == "由负转正" else "turn_neg"},
        "boll": {"enabled": boll_on, "mode": {"突破上轨": "break_upper", "跌破下轨": "break_lower",
                                               "在中轨上方": "above_mid", "在中轨下方": "below_mid"}[boll_mode] if boll_on else "break_upper"},
        "kdj_daily": {
            "K": {"enabled": dk_on and dk_k_on, "op": dk_k_op, "value": dk_k_v},
            "D": {"enabled": dk_on and dk_d_on, "op": dk_d_op, "value": dk_d_v},
            "J": {"enabled": dk_on and dk_j_on, "op": dk_j_op, "value": dk_j_v},
        },
        "kdj_cross": {"enabled": kdj_cross_on, "mode": "golden" if "金叉" in kdj_cross_mode else "dead"},
        "kdj_weekly": {
            "K": {"enabled": wk_on and wk_k_on, "op": wk_k_op, "value": wk_k_v},
            "D": {"enabled": wk_on and wk_d_on, "op": wk_d_op, "value": wk_d_v},
            "J": {"enabled": wk_on and wk_j_on, "op": wk_j_op, "value": wk_j_v},
        },
        "kdj_monthly": {
            "K": {"enabled": mk_on and mk_k_on, "op": mk_k_op, "value": mk_k_v},
            "D": {"enabled": mk_on and mk_d_on, "op": mk_d_op, "value": mk_d_v},
            "J": {"enabled": mk_on and mk_j_on, "op": mk_j_op, "value": mk_j_v},
        },
        "rsi": {"enabled": rsi_on, "period": rsi_period, "op": rsi_op, "value": rsi_val},
        "volume": {"enabled": vol_on, "multiple": vol_mult},
        "vol_price": {"enabled": vp_on, "mode": "up" if vp_mode == "量增价升" else "down"},
        "atr": {"enabled": atr_on, "op": atr_op, "value": atr_val},
        "break_high": {"enabled": bh_on, "n": bh_n},
        "break_low":  {"enabled": bl_on, "n": bl_n},
        "streak": {
            "enabled": streak_on,
            "direction": "up" if "涨" in streak_direction else "down",
            "op": streak_op,
            "days": streak_days,
        },
    }
    use_and = logic_mode.startswith("同时")

    index_code = INDEX_CODE_MAP[index_name]
    with st.spinner(f"正在获取 {index_name} 成分股..."):
        comp_df, comp_err = get_index_components(index_code)
        if comp_df is None:
            st.error(f"获取成分股失败：{comp_err}"); st.stop()

    # 自选股过滤
    if only_watchlist:
        wl = load_watchlist()
        if not wl:
            st.error("自选股为空，请先添加或取消「只在自选股中筛选」。")
            st.stop()
        comp_df = comp_df[comp_df["code"].isin(wl)].reset_index(drop=True)
        st.info(f"已过滤为自选股 {len(comp_df)} 只")

    st.success(f"✅ {index_name} 共 {len(comp_df)} 只股票")
    suspended_set = get_suspended_stocks() if skip_st else set()

    # ===== 第一阶段：技术面 =====
    st.markdown("### 🔍 第一阶段：技术面筛选")
    progress_bar = st.progress(0, text="准备开始...")
    status_text = st.empty()

    tech_passed = []
    all_scanned = []
    failed = 0
    skipped = 0
    fail_reasons = {}

    if use_concurrent:
        st.info(f"并发模式：{num_workers} 线程")
        rows = list(comp_df.iterrows())
        with ThreadPoolExecutor(max_workers=num_workers) as ex:
            futures = {ex.submit(process_one_stock, row["code"], row["name"],
                                  history_days, conditions, use_and, skip_st, suspended_set): row
                       for _, row in rows}
            for i, fut in enumerate(as_completed(futures)):
                res = fut.result()
                progress_bar.progress((i + 1) / len(rows), text=f"进度 {i+1}/{len(rows)}")
                if res["status"] == "skipped":
                    skipped += 1
                elif res["status"] == "failed":
                    failed += 1
                    fail_reasons[res["code"]] = res["error"]
                else:
                    all_scanned.append({"代码": res["code"], "名称": res["name"], **res["snap"]})
                    if res["ok"]:
                        tech_passed.append({"代码": res["code"], "名称": res["name"],
                                            **res["snap"], "命中条件": " | ".join(res["hits"])})
    else:
        for i, row in comp_df.iterrows():
            code, name = row["code"], row["name"]
            status_text.text(f"处理 [{i+1}/{len(comp_df)}] {code} {name}")
            progress_bar.progress((i + 1) / len(comp_df), text=f"进度 {i+1}/{len(comp_df)} — {code} {name}")
            res = process_one_stock(code, name, history_days, conditions, use_and, skip_st, suspended_set)
            if res["status"] == "skipped":
                skipped += 1
            elif res["status"] == "failed":
                failed += 1
                fail_reasons[code] = res["error"]
            else:
                all_scanned.append({"代码": code, "名称": name, **res["snap"]})
                if res["ok"]:
                    tech_passed.append({"代码": code, "名称": name,
                                        **res["snap"], "命中条件": " | ".join(res["hits"])})

    progress_bar.empty()
    status_text.empty()
    st.info(f"第一阶段完成：技术面命中 **{len(tech_passed)}** 只")

    # ===== 第二阶段：基本面 =====
    final_results = []
    if fund_on and tech_passed:
        st.markdown("### 💰 第二阶段：基本面筛选")
        fund_bar = st.progress(0, text="拉取基本面...")
        fund_status = st.empty()
        fund_pass = []
        for i, item in enumerate(tech_passed):
            code, name = item["代码"], item["名称"]
            fund_status.text(f"处理 [{i+1}/{len(tech_passed)}] {code} {name}")
            fund_bar.progress((i + 1) / len(tech_passed), text=f"进度 {i+1}/{len(tech_passed)}")
            fund = get_fundamental_data(code)
            item.update({"PE": fund["PE"], "PB": fund["PB"], "总市值(亿)": fund["总市值(亿)"],
                         "ROE(%)": fund["ROE(%)"], "净利润增速(%)": fund["净利润增速(%)"],
                         "营收增速(%)": fund["营收增速(%)"]})
            ok = True
            if pe_on and fund["PE"] is not None:
                if not (pe_min <= fund["PE"] <= pe_max): ok = False
            if pb_on and fund["PB"] is not None:
                if not (pb_min <= fund["PB"] <= pb_max): ok = False
            if roe_on and fund["ROE(%)"] is not None:
                if fund["ROE(%)"] < roe_min: ok = False
            if mv_on and fund["总市值(亿)"] is not None:
                if fund["总市值(亿)"] < mv_min: ok = False
            if profit_on and fund["净利润增速(%)"] is not None:
                if fund["净利润增速(%)"] < profit_min: ok = False
            if rev_on and fund["营收增速(%)"] is not None:
                if fund["营收增速(%)"] < rev_min: ok = False
            if ok:
                fund_pass.append(item)
        fund_bar.empty()
        fund_status.empty()
        st.info(f"第二阶段完成：基本面通过 **{len(fund_pass)}** 只")
        final_results = fund_pass
    else:
        final_results = tech_passed

    # ===== 评分 & 排序 =====
    for item in final_results:
        hits = item["命中条件"].split(" | ") if item["命中条件"] else []
        fund_dict = {"PE": item.get("PE"), "ROE(%)": item.get("ROE(%)")}
        score, detail = compute_score(hits, fund_dict)
        item["综合评分"] = score
        item["评分明细"] = detail
    final_results = sorted(final_results, key=lambda x: x.get("综合评分", 0), reverse=True)

        # ===== 保存历史 =====
    save_history(index_name, logic_mode, conditions, final_results)
    st.session_state["last_results"] = final_results
    st.session_state["last_all_scanned"] = all_scanned
    st.session_state["last_tech_passed_count"] = len(tech_passed)
    st.session_state["last_failed"] = failed
    st.session_state["last_skipped"] = skipped
    st.session_state["last_fail_reasons"] = fail_reasons
    st.session_state["last_index_name"] = index_name
    st.session_state["last_comp_count"] = len(comp_df)


# ============================================================
# 历史筛选记录（页面底部）
# ============================================================
st.divider()
st.subheader("📚 历史筛选记录")

hist_df = load_history(limit=30)
if hist_df.empty:
    st.caption("暂无历史记录。每次筛选后自动保存。")
else:
    hist_display = hist_df.copy()
    hist_display["条件摘要"] = hist_display["conditions"].apply(summarize_conditions)
    hist_display = hist_display.rename(columns={
        "id": "ID", "run_time": "时间", "index_name": "股票池",
        "logic_mode": "逻辑", "result_count": "结果数",
    })
    hist_display = hist_display[["ID", "时间", "股票池", "逻辑", "结果数", "条件摘要"]]

    st.dataframe(
        hist_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ID": st.column_config.NumberColumn("ID", width="small", format="%d"),
            "时间": st.column_config.TextColumn("时间", width="medium"),
            "股票池": st.column_config.TextColumn("股票池", width="small"),
            "逻辑": st.column_config.TextColumn("逻辑", width="small"),
            "结果数": st.column_config.NumberColumn("结果数", width="small"),
            "条件摘要": st.column_config.TextColumn("条件摘要", width="large"),
        },
    )
    
    # ---------- 删除功能 ----------
    del_col1, del_col2, del_col3 = st.columns([2, 1, 1])
    with del_col1:
        del_options = ["不删除"] + [
            f"#{r['id']} - {r['run_time']} ({r['index_name']}, {r['result_count']}只)"
            for _, r in hist_df.iterrows()
        ]
        del_target = st.selectbox("选择要删除的记录", del_options, key="del_hist_sel")
    with del_col2:
        st.write("")
        st.write("")
        if st.button("🗑️ 删除选中", key="btn_del_one", use_container_width=True):
            if del_target != "不删除":
                try:
                    hid = int(del_target.split(" ")[0].replace("#", ""))
                    if delete_history(hid):
                        st.success(f"已删除记录 #{hid}")
                        st.rerun()
                    else:
                        st.error("删除失败，请重试。")
                except Exception as e:
                    st.error(f"删除出错：{e}")
    with del_col3:
        st.write("")
        st.write("")
        if st.button("⚠️ 清空全部", key="btn_del_all", use_container_width=True):
            if clear_all_history():
                st.success("已清空全部历史记录")
                st.rerun()
            else:
                st.error("清空失败，请重试。")

    hist_id = st.selectbox(
        "查看某次详情",
        ["不查看"] + [
            f"#{r['id']} - {r['run_time']} | {r['index_name']} | {r['result_count']}只"
            for _, r in hist_df.iterrows()
        ],
        key="hist_sel",
    )
    if hist_id != "不查看":
        try:
            hid = int(hist_id.split(" ")[0].replace("#", ""))
            detail = load_history_detail(hid)
            if detail and detail["results"]:
                idx_name = detail.get("index_name") or "未知指数"
                st.markdown(f"**{detail['run_time']} | {idx_name}**")
                hdf = pd.DataFrame(detail["results"])
                if "综合评分" in hdf.columns:
                    hdf = hdf.sort_values("综合评分", ascending=False)
                st.dataframe(hdf, use_container_width=True, height=400)
            else:
                st.info("该次筛选没有结果。")
        except Exception as e:
            st.error(f"读取历史失败：{e}")

st.caption("💡 数据来自东方财富、新浪财经、乐咕乐股（AKShare）。缓存目录：data/。")



# ============================================================
# 回测分析（追加在文件末尾）
# ============================================================
st.divider()


# ============================================================
# 回测分析（升级版：手续费 + 基准对比 + 样本量提示）
# ============================================================
@st.cache_data(ttl=3600, show_spinner=False)
def get_index_hist(index_code="000300", days=600):
    """获取指数历史行情，用于基准对比"""
    end_date = pd.Timestamp.today().strftime("%Y%m%d")
    start_date = (pd.Timestamp.today() - pd.Timedelta(days=days * 2)).strftime("%Y%m%d")
    try:
        df = ak.index_zh_a_hist(symbol=index_code, period="daily",
                                start_date=start_date, end_date=end_date)
        if df is None or df.empty:
            return None
        df = df.rename(columns={"日期": "date", "收盘": "close"})
        df = df[["date", "close"]].copy()
        df["date"] = pd.to_datetime(df["date"])
        return df.sort_values("date").reset_index(drop=True)
    except Exception:
        return None


def calc_benchmark_return(bench_df, screen_date, hold_days):
    """计算沪深300同期涨幅"""
    if bench_df is None or bench_df.empty:
        return {}
    before = bench_df[bench_df["date"] <= screen_date]
    after = bench_df[bench_df["date"] > screen_date].reset_index(drop=True)
    if before.empty:
        return {}
    buy_price = before["close"].iloc[-1]
    result = {}
    for n in hold_days:
        if len(after) >= n:
            sell_price = after["close"].iloc[n - 1]
            result[n] = round((sell_price / buy_price - 1) * 100, 2)
    return result


st.divider()
st.subheader("🧪 回测分析")
st.caption("假设筛选日收盘价买入，持有 N 个交易日后按收盘价卖出。已扣除交易成本。")

# ---------- 交易成本设置 ----------
with st.expander("⚙️ 交易成本设置（影响净收益计算）", expanded=False):
    st.markdown("""
**A股单次交易成本说明**：
- 佣金：万分之2.5（双向），最低5元
- 印花税：千分之1（**仅卖出**）
- 过户费：万分之0.1（双向）
- 滑点：买入/卖出各约 0.05%~0.1%

**简化后的总双边成本，默认 0.3%**，如果做短线（1日、5日）建议适当调高。
    """)
    cost_pct = st.slider(
        "单次交易总成本（%）",
        min_value=0.0, max_value=1.0, value=0.3, step=0.05,
        key="bt_cost",
        help="买入+卖出的总成本。0.3% 是常见估算，包含佣金、印花税、过户费、滑点。",
    )

bt_hist_df = load_history(limit=50)

if bt_hist_df.empty:
    st.info("暂无历史筛选记录。请先执行至少一次筛选。")
else:
    options = ["不选择"] + [
        f"#{r['id']} - {r['run_time']} | {r['index_name']} | {r['result_count']}只"
        for _, r in bt_hist_df.iterrows()
        if r["result_count"] > 0
    ]

    if len(options) == 1:
        st.info("暂无包含结果的历史记录。")
    else:
        sel_hist = st.selectbox("选择要回测的历史记录", options, key="bt_hist_sel")

        bc1, bc2 = st.columns([3, 1])
        with bc1:
            hold_days_input = st.text_input(
                "持有天数（逗号分隔，最多 5 个）",
                "1,5,10,20,40",
                key="bt_hold",
            )
        with bc2:
            st.write("")
            st.write("")
            run_bt = st.button("🚀 执行回测", type="primary", key="bt_run", use_container_width=True)

        if run_bt and sel_hist != "不选择":
            try:
                hold_days = [int(x.strip()) for x in hold_days_input.split(",") if x.strip().isdigit()]
                hold_days = hold_days[:5]
            except Exception:
                hold_days = [1, 5, 10, 20]

            try:
                hid = int(sel_hist.split(" ")[0].replace("#", ""))
            except Exception:
                st.error("无法解析历史记录 ID。")
                st.stop()

            detail = load_history_detail(hid)
            if not detail or not detail["results"]:
                st.warning("该次记录没有可回测的股票。")
            else:
                screen_time = detail["run_time"]
                screen_date = pd.to_datetime(screen_time).normalize()
                results = detail["results"]

                st.markdown(
                    f"**回测基准日**：{screen_date.strftime('%Y-%m-%d')}　"
                    f"（筛选执行时间：{screen_time}）　"
                    f"**股票数**：{len(results)} 只"
                )

                # ---------- 样本量提示 ----------
                n_samples = len(results)
                if n_samples < 5:
                    st.error(f"⚠️ 样本量仅 **{n_samples}** 只，统计意义极弱。建议至少 10 只以上再做判断。")
                elif n_samples < 10:
                    st.warning(f"⚠️ 样本量 **{n_samples}** 只，勉强可用。建议 ≥10 只。")
                elif n_samples < 30:
                    st.info(f"ℹ️ 样本量 **{n_samples}** 只，达到基本可用标准。")
                else:
                    st.success(f"✅ 样本量 **{n_samples}** 只，统计意义较好。")

                # ---------- 计算基准 ----------
                with st.spinner("正在拉取沪深300基准数据..."):
                    bench_df = get_index_hist("000300", days=600)
                    bench_returns = calc_benchmark_return(bench_df, screen_date, hold_days)

                # ---------- 逐只回测 ----------
                progress = st.progress(0, text="回测中...")
                bt_rows = []
                for i, item in enumerate(results):
                    progress.progress(
                        (i + 1) / len(results),
                        text=f"处理 [{i+1}/{len(results)}] {item.get('代码','')} {item.get('名称','')}",
                    )
                    code = str(item.get("代码", "")).zfill(6)
                    name = item.get("名称", "")
                    buy_price = item.get("最新价")
                    if buy_price is None or buy_price == 0:
                        continue

                    kline, _ = load_or_update_kline(code, days=500)
                    if kline is None or len(kline) == 0:
                        continue

                    future = kline[kline["date"] > screen_date].reset_index(drop=True)

                    row = {
                        "代码": code,
                        "名称": name,
                        "综合评分": item.get("综合评分", 0),
                        "筛选日价格": buy_price,
                    }
                    for n in hold_days:
                        if len(future) >= n:
                            sell_price = future["close"].iloc[n - 1]
                            raw_ret = (sell_price / buy_price - 1) * 100
                            net_ret = raw_ret - cost_pct
                            row[f"{n}日毛收益%"] = round(raw_ret, 2)
                            row[f"{n}日净收益%"] = round(net_ret, 2)
                        else:
                            row[f"{n}日毛收益%"] = None
                            row[f"{n}日净收益%"] = None

                    if len(future) > 0:
                        max_high = future["high"].max()
                        min_low = future["low"].min()
                        row["最大涨幅%"] = round((max_high / buy_price - 1) * 100, 2)
                        row["最大回撤%"] = round((min_low / buy_price - 1) * 100, 2)
                    else:
                        row["最大涨幅%"] = None
                        row["最大回撤%"] = None

                    bt_rows.append(row)

                progress.empty()

                if not bt_rows:
                    st.warning("没有可回测的数据。")
                else:
                    bt_df = pd.DataFrame(bt_rows)

                    # ---------- 汇总统计 ----------
                    st.markdown("### 📊 汇总统计")
                    stats = []
                    for n in hold_days:
                        col = f"{n}日净收益%"
                        if col not in bt_df.columns:
                            continue
                        valid = bt_df[col].dropna()
                        if len(valid) == 0:
                            continue
                        wins = valid[valid > 0]
                        losses = valid[valid < 0]
                        pnl_ratio = None
                        if len(wins) > 0 and len(losses) > 0:
                            pnl_ratio = round(wins.mean() / abs(losses.mean()), 2)

                        bench_ret = bench_returns.get(n)
                        excess = None
                        if bench_ret is not None:
                            excess = round(valid.mean() - bench_ret, 2)

                        stats.append({
                            "持有期": f"{n}日",
                            "样本数": len(valid),
                            "平均净收益%": round(valid.mean(), 2),
                            "胜率%": round((valid > 0).sum() / len(valid) * 100, 2),
                            "最大盈利%": round(valid.max(), 2),
                            "最大亏损%": round(valid.min(), 2),
                            "盈亏比": pnl_ratio,
                            "沪深300同期%": bench_ret,
                            "超额收益%": excess,
                        })

                    if stats:
                        stats_df = pd.DataFrame(stats)
                        st.dataframe(
                            stats_df,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "持有期": st.column_config.TextColumn("持有期", width="small"),
                                "样本数": st.column_config.NumberColumn("样本数", width="small"),
                                "平均净收益%": st.column_config.NumberColumn("平均净收益%", format="%.2f"),
                                "胜率%": st.column_config.NumberColumn("胜率%", format="%.2f"),
                                "最大盈利%": st.column_config.NumberColumn("最大盈利%", format="%.2f"),
                                "最大亏损%": st.column_config.NumberColumn("最大亏损%", format="%.2f"),
                                "盈亏比": st.column_config.NumberColumn("盈亏比", format="%.2f"),
                                "沪深300同期%": st.column_config.NumberColumn("沪深300同期%", format="%.2f"),
                                "超额收益%": st.column_config.NumberColumn("超额收益%", format="%.2f"),
                            },
                        )

                        st.caption(
                            "💡 **盈亏比**：平均盈利 ÷ 平均亏损。>1 表示赚的比亏的多。\n"
                            "💡 **超额收益**：策略平均收益 − 沪深300同期涨幅。>0 表示跑赢大盘。"
                        )

                    # ---------- 每只股票明细 ----------
                    st.markdown("### 📋 每只股票明细")

                    # 明细列重排：代码、名称、评分、筛选日价格、各持有期净收益、最大涨跌、毛收益（隐藏）
                    net_cols = [f"{n}日净收益%" for n in hold_days if f"{n}日净收益%" in bt_df.columns]
                    show_cols = ["代码", "名称", "综合评分", "筛选日价格"] + net_cols + ["最大涨幅%", "最大回撤%"]
                    show_cols = [c for c in show_cols if c in bt_df.columns]
                    bt_display = bt_df[show_cols].sort_values("综合评分", ascending=False)

                    st.dataframe(bt_display, use_container_width=True, height=420)

                    # ---------- 下载 ----------
                    csv_bt = bt_df.to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        "📥 下载回测明细 CSV",
                        data=csv_bt,
                        file_name=f"回测明细_{screen_date.strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                    )

                    st.caption(
                        f"💡 卖出价用的是筛选日之后第 N 个交易日的收盘价，"
                        f"已扣除 {cost_pct}% 的双边交易成本。未考虑停牌、涨跌停无法成交等情况。"
                    )
                    

# ============================================================
# 结果展示（独立块：即使点击表格行触发 rerun，也能稳定显示）
# ============================================================
if "last_results" in st.session_state:
    final_results = st.session_state["last_results"]
    all_scanned = st.session_state.get("last_all_scanned", [])
    tech_passed_count = st.session_state.get("last_tech_passed_count", 0)
    failed = st.session_state.get("last_failed", 0)
    skipped = st.session_state.get("last_skipped", 0)
    fail_reasons = st.session_state.get("last_fail_reasons", {})
    index_name_display = st.session_state.get("last_index_name", "")
    comp_count = st.session_state.get("last_comp_count", 0)

    st.divider()
    st.subheader("🎉 最终筛选结果")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("扫描股票数", comp_count)
    c2.metric("技术面命中", tech_passed_count)
    c3.metric("最终结果", len(final_results))
    c4.metric("已过滤", skipped)

    if failed > 0:
        with st.expander(f"⚠️ {failed} 只拉取失败"):
            for k, v in list(fail_reasons.items())[:10]:
                st.write(f"**{k}**: {v}")

    if final_results:
        df = pd.DataFrame(final_results)
        col_order = ["代码", "名称", "综合评分", "最新价", "涨跌幅(%)",
                     "PE", "PB", "ROE(%)", "总市值(亿)", "净利润增速(%)", "营收增速(%)",
                     "日K", "日D", "日J", "周K", "周D", "周J", "月K", "月D", "月J",
                     "RSI6", "ATR%", "命中条件", "评分明细"]
        display_cols = [c for c in col_order if c in df.columns]
        df_display = df[display_cols]

        sort_col1, sort_col2 = st.columns([3, 1])
        with sort_col2:
            sort_by = st.selectbox(
                "排序依据",
                ["综合评分", "涨跌幅(%)", "最新价"] + [c for c in ["周K", "月K", "RSI6", "PE"] if c in df_display.columns],
                index=0, key="sort_by",
            )
            sort_asc = st.checkbox("升序", value=False, key="sort_asc")
        df_display = df_display.sort_values(sort_by, ascending=sort_asc).reset_index(drop=True)

        exp1, exp2 = st.columns(2)
        with exp1:
            csv_data = df.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "📥 下载 CSV", data=csv_data,
                file_name=f"筛选结果_{index_name_display}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True,
            )
        with exp2:
            if HAS_OPENPYXL:
                xlsx_data = export_to_excel(df)
                st.download_button(
                    "📊 下载 Excel（带格式）", data=xlsx_data,
                    file_name=f"筛选结果_{index_name_display}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.caption("安装 openpyxl 后可导出 Excel")

        st.caption("💡 点击表格中的任意一行，右侧会显示该股票的K线图和详情。")
        col_left, col_right = st.columns([3, 2])

        with col_left:
            sel = st.dataframe(
                df_display, use_container_width=True, height=520,
                on_select="rerun", selection_mode="single-row", key="result_table",
            )

        with col_right:
            selected_rows = []
            try:
                if sel is not None and hasattr(sel, "selection"):
                    selected_rows = sel.selection.get("rows", []) if isinstance(sel.selection, dict) else sel.selection.rows
            except Exception:
                selected_rows = []
            if selected_rows:
                row_idx = selected_rows[0]
                sc = str(df_display.iloc[row_idx]["代码"]).zfill(6)
                sn = df_display.iloc[row_idx]["名称"]
                st.markdown(f"### 📌 {sc} {sn}")
                detail_df, _ = load_or_update_kline(sc, days=min(history_days, 250))
                if detail_df is not None and len(detail_df) > 20:
                    chart_df = detail_df.set_index("date")[["close"]].copy()
                    chart_df.columns = ["收盘价"]
                    st.line_chart(chart_df, height=200)
                    try:
                        ddf = wrap(detail_df.copy())
                        _ = ddf.get("kdj"); _ = ddf.get("macd"); _ = ddf.get("atr")
                        m1 = st.columns(3)
                        m1[0].metric("最新价", round(detail_df["close"].iloc[-1], 2))
                        m1[1].metric("日K", round(ddf["kdjk"].iloc[-1], 2))
                        m1[2].metric("日D", round(ddf["kdjd"].iloc[-1], 2))
                        m2 = st.columns(3)
                        m2[0].metric("日J", round(ddf["kdjj"].iloc[-1], 2))
                        m2[1].metric("MACD", round(ddf["macd"].iloc[-1], 3))
                        m2[2].metric("ATR%", round(ddf["atr"].iloc[-1] / detail_df["close"].iloc[-1] * 100, 2))
                    except Exception:
                        pass
                    recent = detail_df.tail(5)[["date", "open", "high", "low", "close", "volume"]].copy()
                    recent["date"] = recent["date"].dt.strftime("%Y-%m-%d")
                    recent.columns = ["日期", "开盘", "最高", "最低", "收盘", "成交量"]
                    st.dataframe(recent, use_container_width=True, hide_index=True)
                else:
                    st.warning("无法加载K线数据。")
            else:
                st.info("👈 点击左侧表格行查看详情。")
    else:
        st.warning("⚠️ 没有股票满足条件。")
        if all_scanned:
            st.subheader("📊 已扫描快照（前20只）")
            st.dataframe(pd.DataFrame(all_scanned).head(20), use_container_width=True)
