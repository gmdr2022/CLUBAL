# weather.py
# SAL - Weather module (online + cache fallback)
# Cache gravável (preferência): %LOCALAPPDATA%/SAL_SESI_Agenda_Live/package/weather_cache.json
# ✅ Refinamentos:
# - Rotação/arquivo de caches antigos em package/cache_old/
# - Limite por quantidade + idade (auto-limpeza)
# - Escrita atômica
# SSL robusto:
# 1) truststore (usa certificados do Windows)
# 2) certifi (fallback)
# 3) default SSL
# Nunca trava UI – sempre cai pro cache se falhar (e sal.py chama em thread)

from __future__ import annotations

import json
import os
import ssl
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple
import urllib.request
import urllib.error


@dataclass
class WeatherResult:
    ok: bool
    temp_c: Optional[int]
    today_label: str
    tomorrow_label: str
    symbol_code: Optional[str]
    source: str                # "online" | "cache"
    cache_ts: Optional[int]


# -------------------------
# Cache policy (refinamentos)
# -------------------------

CACHE_ARCHIVE_DIRNAME = "cache_old"
CACHE_ARCHIVE_KEEP = 10                 # mantém últimos N caches antigos
CACHE_ARCHIVE_MAX_AGE_DAYS = 7          # apaga cache_old com mais de N dias
CACHE_STALE_WARN_SECONDS = 6 * 3600     # (opcional) definir "velho" > 6h (UI já mostra horário)


def _safe_int(x: Any) -> Optional[int]:
    try:
        return int(round(float(x)))
    except Exception:
        return None


def _safe_mkdir(path: str) -> None:
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass


def _read_json(path: str) -> Optional[Dict[str, Any]]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _write_json_atomic(path: str, data: Dict[str, Any]) -> None:
    d = os.path.dirname(path)
    _safe_mkdir(d)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _cache_root(app_dir: str) -> str:
    """
    Raiz gravável para cache.
    Preferência: %LOCALAPPDATA%/SAL_SESI_Agenda_Live/package
    Fallback: app_dir/package (melhor esforço)
    """
    base = os.environ.get("LOCALAPPDATA")
    if base:
        p = os.path.join(base, "SAL_SESI_Agenda_Live", "package")
        _safe_mkdir(p)
        return p
    p = os.path.join(app_dir, "package")
    _safe_mkdir(p)
    return p


def _cache_paths(app_dir: str) -> Tuple[str, str]:
    root = _cache_root(app_dir)
    current = os.path.join(root, "weather_cache.json")
    archive = os.path.join(root, CACHE_ARCHIVE_DIRNAME)
    _safe_mkdir(archive)
    return current, archive


def _archive_existing_cache(current_path: str, archive_dir: str, logger=None) -> None:
    """
    Move o cache atual para cache_old com timestamp no nome, antes de gravar um novo.
    Isso separa “cache usual” do histórico e facilita limpeza.
    """
    try:
        if not os.path.exists(current_path):
            return

        cached = _read_json(current_path) or {}
        ts = cached.get("ts")
        if isinstance(ts, int) and ts > 0:
            stamp = time.strftime("%Y%m%d_%H%M%S", time.localtime(ts))
        else:
            stamp = time.strftime("%Y%m%d_%H%M%S")

        dst = os.path.join(archive_dir, f"weather_cache_{stamp}.json")

        # evita sobrescrever se já existir
        if os.path.exists(dst):
            dst = os.path.join(archive_dir, f"weather_cache_{stamp}_{int(time.time())}.json")

        os.replace(current_path, dst)
        if logger:
            logger(f"[WEATHER] Archived old cache -> {dst}")

    except Exception as e:
        if logger:
            logger(f"[WEATHER] Archive cache error {type(e).__name__}: {e}")


def _cleanup_cache_archive(archive_dir: str, logger=None) -> None:
    """
    Limpa cache_old por:
    - idade (CACHE_ARCHIVE_MAX_AGE_DAYS)
    - quantidade (CACHE_ARCHIVE_KEEP)
    """
    try:
        now = time.time()
        max_age = CACHE_ARCHIVE_MAX_AGE_DAYS * 86400

        files = []
        for name in os.listdir(archive_dir):
            p = os.path.join(archive_dir, name)
            if not os.path.isfile(p):
                continue
            try:
                st = os.stat(p)
            except Exception:
                continue

            if max_age > 0 and (now - st.st_mtime) > max_age:
                try:
                    os.remove(p)
                    if logger:
                        logger(f"[WEATHER] Pruned old archive (age) {p}")
                except Exception:
                    pass
                continue

            files.append((st.st_mtime, p))

        files.sort(reverse=True)  # newest first
        for _mtime, p in files[CACHE_ARCHIVE_KEEP:]:
            try:
                os.remove(p)
                if logger:
                    logger(f"[WEATHER] Pruned old archive (count) {p}")
            except Exception:
                pass

    except Exception as e:
        if logger:
            logger(f"[WEATHER] Cleanup archive error {type(e).__name__}: {e}")


def housekeeping(app_dir: str, logger=None) -> None:
    """
    Pode ser chamado no boot e 1x/dia.
    """
    current, archive = _cache_paths(app_dir)
    _cleanup_cache_archive(archive, logger=logger)

    # também remove tmp velho se existir
    try:
        tmp = current + ".tmp"
        if os.path.exists(tmp):
            os.remove(tmp)
    except Exception:
        pass


# -------------------------
# SSL / HTTP
# -------------------------

def _ssl_context_best_effort() -> ssl.SSLContext:
    """
    Ordem:
    1) truststore → usa certificados do Windows (ideal em rede corporativa)
    2) certifi
    3) default Python
    """
    try:
        import truststore  # type: ignore
        truststore.inject_into_ssl()
        return ssl.create_default_context()
    except Exception:
        pass

    try:
        import certifi  # type: ignore
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        pass

    return ssl.create_default_context()


def _build_opener(logger=None) -> urllib.request.OpenerDirector:
    proxies: Dict[str, str] = {}

    try:
        p_env = urllib.request.getproxies() or {}
        proxies.update({k.lower(): v for k, v in p_env.items() if v})
    except Exception:
        pass

    try:
        p_reg = urllib.request.getproxies_registry() or {}
        proxies.update({k.lower(): v for k, v in p_reg.items() if v})
    except Exception:
        pass

    if logger:
        logger(f"[WEATHER] Proxies detectados: {proxies if proxies else 'nenhum'}")

    proxy_handler = urllib.request.ProxyHandler(proxies) if proxies else urllib.request.ProxyHandler({})
    https_handler = urllib.request.HTTPSHandler(context=_ssl_context_best_effort())
    return urllib.request.build_opener(proxy_handler, https_handler)


def _http_get_json(url: str, user_agent: str, timeout: int = 6, logger=None) -> Dict[str, Any]:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": user_agent, "Accept": "application/json"},
        method="GET",
    )

    opener = _build_opener(logger=logger)

    # retry simples (ajuda oscilação)
    last_exc: Optional[Exception] = None
    for attempt in (1, 2):
        try:
            if logger:
                logger(f"[WEATHER] HTTP attempt {attempt} timeout={timeout}s")
            with opener.open(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
                if logger:
                    status = getattr(resp, "status", None) or getattr(resp, "code", "?")
                    logger(f"[WEATHER] HTTP {status} len={len(raw)}")
                return json.loads(raw)

        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read(200).decode("utf-8", errors="replace")
            except Exception:
                pass
            if logger:
                logger(f"[WEATHER] HTTPError {e.code} {e.reason} body='{body}'")
            raise

        except urllib.error.URLError as e:
            last_exc = e
            if logger:
                logger(f"[WEATHER] URLError reason={repr(getattr(e, 'reason', e))}")
            if attempt == 1:
                time.sleep(0.4)
                continue
            raise

        except Exception as e:
            last_exc = e
            if logger:
                logger(f"[WEATHER] Exception {type(e).__name__}: {e}")
            if attempt == 1:
                time.sleep(0.4)
                continue
            raise

    raise last_exc if last_exc else RuntimeError("HTTP failed")


# -------------------------
# Parsing / Labels
# -------------------------

def _pick_period(now_hour: int) -> str:
    if 6 <= now_hour <= 11:
        return "manhã"
    if 12 <= now_hour <= 17:
        return "tarde"
    return "noite"


def _minmax_tomorrow(timeseries: list) -> Tuple[Optional[int], Optional[int]]:
    tmin = None
    tmax = None
    for item in timeseries:
        inst = item.get("data", {}).get("instant", {}).get("details", {})
        temp = _safe_int(inst.get("air_temperature"))
        if temp is None:
            continue
        tmin = temp if tmin is None else min(tmin, temp)
        tmax = temp if tmax is None else max(tmax, temp)
    return tmin, tmax


def _extract_summary(payload: Dict[str, Any], now_hour: int) -> WeatherResult:
    props = payload.get("properties", {})
    timeseries = props.get("timeseries", [])

    temp_now = None
    symbol_code = None

    if timeseries:
        first = timeseries[0].get("data", {})
        temp_now = _safe_int(first.get("instant", {}).get("details", {}).get("air_temperature"))
        n1 = first.get("next_1_hours", {}).get("summary", {}).get("symbol_code")
        n6 = first.get("next_6_hours", {}).get("summary", {}).get("symbol_code")
        symbol_code = n1 or n6

    period = _pick_period(now_hour)

    def humanize(sym: Optional[str]) -> str:
        if not sym:
            return "Sem dados"
        s = sym.lower()
        if "thunder" in s:
            return "Tempestade"
        if "snow" in s:
            return "Neve"
        if "rain" in s or "sleet" in s:
            if "heavyrain" in s or "rainshowersandthunder" in s or "heavyrainshowers" in s:
                return "Chuva forte"
            if "lightrain" in s or "lightrainshowers" in s:
                return "Chuva fraca"
            return "Chuva"
        if "cloudy" in s:
            return "Parcialmente nublado" if "partly" in s else "Nublado"
        if "clearsky" in s or "fair" in s:
            return "Céu limpo"
        return "Tempo instável"

    today_label = f"Hoje ({period}): {humanize(symbol_code)}"

    tmin, tmax = _minmax_tomorrow(timeseries)
    tomorrow_label = f"Amanhã: {tmin}–{tmax}°C" if (tmin is not None and tmax is not None) else "Amanhã: —"

    return WeatherResult(
        ok=True,
        temp_c=temp_now,
        today_label=today_label,
        tomorrow_label=tomorrow_label,
        symbol_code=symbol_code,
        source="online",
        cache_ts=int(time.time()),
    )


# -------------------------
# Public API
# -------------------------

def get_weather(
    city_label: str,
    lat: float,
    lon: float,
    app_dir: str,
    user_agent: str = "SAL-SESIAgendaLive/2.0 (contact: gui@sesi.local)",
    logger=None,
) -> WeatherResult:
    """
    Returns WeatherResult.
    - Tries online from met.no
    - Falls back to cache
    - ✅ rotates old cache to cache_old/ (with limits)
    """
    current_cache_path, archive_dir = _cache_paths(app_dir)
    now_hour = time.localtime().tm_hour

    url = (
        "https://api.met.no/weatherapi/locationforecast/2.0/compact"
        f"?lat={lat:.4f}&lon={lon:.4f}"
    )

    try:
        if logger:
            logger(f"[WEATHER] Fetch start url={url}")

        payload = _http_get_json(url, user_agent=user_agent, timeout=6, logger=logger)
        res = _extract_summary(payload, now_hour=now_hour)

        # ✅ antes de escrever o novo cache, arquiva o atual
        _archive_existing_cache(current_cache_path, archive_dir, logger=logger)

        _write_json_atomic(current_cache_path, {"ts": res.cache_ts, "payload": payload, "city": city_label})

        # ✅ limpa histórico velho (auto-manutenção)
        _cleanup_cache_archive(archive_dir, logger=logger)

        if logger:
            logger(f"[WEATHER] ONLINE ok temp={res.temp_c} sym={res.symbol_code} cache_path={current_cache_path}")

        return res

    except Exception:
        cached = _read_json(current_cache_path)
        if cached and isinstance(cached.get("payload"), dict):
            payload = cached["payload"]
            res = _extract_summary(payload, now_hour=now_hour)
            res.source = "cache"
            res.cache_ts = cached.get("ts")
            res.ok = True

            if logger:
                logger(f"[WEATHER] FALLBACK cache ok ts={res.cache_ts} cache_path={current_cache_path}")

            return res

        if logger:
            logger(f"[WEATHER] FAIL no cache available cache_path={current_cache_path}")

        return WeatherResult(
            ok=False,
            temp_c=None,
            today_label="Sem dados",
            tomorrow_label="Amanhã: —",
            symbol_code=None,
            source="cache",
            cache_ts=None,
        )
