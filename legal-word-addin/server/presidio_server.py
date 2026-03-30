#!/usr/bin/env python3
"""
HTTP sidecar: POST /scan { "text": "..." } -> { "items": [{ "original", "category" }] }
GET /health -> 200 when the process is up (engine loads on first /scan).

Run: PYTHONPATH=. python3 presidio_server.py
Or via npm / proxy auto-spawn with PRESIDIO_PORT.

Env:
  PRESIDIO_SCORE_THRESHOLD — default 0.2 (lower = more recall).
  PRESIDIO_LANGS — e.g. "en,de" to restrict passes when multiple SpaCy models are installed.
  Install de_core_news_sm / fr_core_news_sm / it_core_news_sm for Swiss DE/FR/IT NER (npm run presidio:install).
"""

from __future__ import annotations

import json
import os
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer

PORT = int(os.environ.get("PRESIDIO_PORT", "3549"))
# Lower = more recall (more false positives). 0.35 misses many valid spans; 0.2 is a better default for legal redaction.
DEFAULT_SCORE_THRESHOLD = float(os.environ.get("PRESIDIO_SCORE_THRESHOLD", "0.2"))
# Comma-separated subset of installed languages, e.g. "en,de". Empty = use all SpaCy models found below.
_PRESIDIO_LANGS_ENV = os.environ.get("PRESIDIO_LANGS", "").strip()

# Presidio entity types -> add-in categories (see redact-fakes.js / office-helpers redactItems).
# Covers global + regional recognizers and clinical NER types from:
# https://microsoft.github.io/presidio/supported_entities/
# SpaCy/Stanza-style labels often merged into Presidio results:
ENTITY_TO_CATEGORY = {
    # --- People & sensitive groups ---
    "PERSON": "person",
    "NRP": "person",
    "NORP": "person",
    # --- Organizations (incl. business / entity registraton numbers as org-like IDs) ---
    "ORGANIZATION": "company",
    "ORG": "company",
    "AU_ABN": "company",
    "AU_ACN": "company",
    "SG_UEN": "company",
    "KR_BRN": "company",
    # --- Locations & postal ---
    "LOCATION": "address",
    "GPE": "address",
    "LOC": "address",
    "FAC": "address",
    "UK_POSTCODE": "address",
    "US_ZIP": "address",
    # --- Contact ---
    "EMAIL_ADDRESS": "email",
    "PHONE_NUMBER": "phone",
    # --- Financial / payment rails ---
    "IBAN_CODE": "bank_account",
    "CREDIT_CARD": "bank_account",
    "US_BANK_NUMBER": "bank_account",
    "US_ROUTING_NUMBER": "bank_account",
    "CRYPTO": "bank_account",
    # --- Government & tax IDs, devices, clinical PHI-style spans ---
    "IP_ADDRESS": "id_number",
    "MAC_ADDRESS": "id_number",
    "URL": "id_number",
    "MEDICAL_LICENSE": "id_number",
    "US_SSN": "id_number",
    "US_PASSPORT": "id_number",
    "US_DRIVER_LICENSE": "id_number",
    "US_ITIN": "id_number",
    "US_MBI": "id_number",
    "US_NPI": "id_number",
    "UK_NHS": "id_number",
    "UK_NINO": "id_number",
    "UK_PASSPORT": "id_number",
    "UK_VEHICLE_REGISTRATION": "id_number",
    "ES_NIF": "id_number",
    "ES_NIE": "id_number",
    "IT_FISCAL_CODE": "id_number",
    "IT_DRIVER_LICENSE": "id_number",
    "IT_VAT_CODE": "id_number",
    "IT_PASSPORT": "id_number",
    "IT_IDENTITY_CARD": "id_number",
    "PL_PESEL": "id_number",
    "SG_NRIC_FIN": "id_number",
    "AU_TFN": "id_number",
    "AU_MEDICARE": "id_number",
    "IN_PAN": "id_number",
    "IN_AADHAAR": "id_number",
    "IN_VEHICLE_REGISTRATION": "id_number",
    "IN_VOTER": "id_number",
    "IN_PASSPORT": "id_number",
    "IN_GSTIN": "id_number",
    "FI_PERSONAL_IDENTITY_CODE": "id_number",
    "KR_DRIVER_LICENSE": "id_number",
    "KR_FRN": "id_number",
    "KR_PASSPORT": "id_number",
    "KR_RRN": "id_number",
    "NG_NIN": "id_number",
    "NG_VEHICLE_REGISTRATION": "id_number",
    "TH_TNIN": "id_number",
    # Swiss / CH — custom PatternRecognizers (UID, social ins., SWIFT); IBAN still via IBAN_CODE
    "CH_UID": "company",
    "CH_AHV_NUMBER": "id_number",
    "SWIFT_CODE": "bank_account",
    "MEDICAL_DISEASE_DISORDER": "id_number",
    "MEDICAL_MEDICATION": "id_number",
    "MEDICAL_THERAPEUTIC_PROCEDURE": "id_number",
    "MEDICAL_CLINICAL_EVENT": "id_number",
    "MEDICAL_BIOLOGICAL_ATTRIBUTE": "id_number",
    "MEDICAL_BIOLOGICAL_STRUCTURE": "id_number",
    "MEDICAL_FAMILY_HISTORY": "id_number",
    "MEDICAL_HISTORY": "id_number",
}

# Types we still skip so Presidio behavior matches the LLM anonymize prompt (omit generic document dates/times).
SKIP_TYPES = frozenset({"DATE_TIME"})

DEFAULT_CATEGORY_UNKNOWN_ENTITY = "id_number"

_engine = None
_custom_recognizers_registered = False
# Languages Presidio will scan (from SpaCy models + env filter)
_ENGINE_ANALYSIS_LANGS: list[str] = ["en"]


def _nlp_engine_and_supported_languages() -> tuple[object | None, list[str]]:
    """Build multi-language SpaCy engine when models are installed (critical for DE/FR/IT Swiss documents)."""
    try:
        import spacy
        from presidio_analyzer.nlp_engine import NlpEngineProvider
    except ImportError:
        return None, []

    pairs = [
        ("en", "en_core_web_sm"),
        ("de", "de_core_news_sm"),
        ("fr", "fr_core_news_sm"),
        ("it", "it_core_news_sm"),
    ]
    models: list[dict[str, str]] = []
    for code, name in pairs:
        try:
            spacy.load(name)
            models.append({"lang_code": code, "model_name": name})
        except OSError:
            pass
    if not models:
        return None, []
    configuration = {"nlp_engine_name": "spacy", "models": models}
    provider = NlpEngineProvider(nlp_configuration=configuration)
    nlp_engine = provider.create_engine()
    langs = [m["lang_code"] for m in models]
    return nlp_engine, langs


def _filter_analysis_languages(all_langs: list[str]) -> list[str]:
    if not _PRESIDIO_LANGS_ENV:
        return list(all_langs)
    requested = {x.strip().lower() for x in _PRESIDIO_LANGS_ENV.split(",") if x.strip()}
    picked = [L for L in all_langs if L.lower() in requested]
    return picked if picked else [all_langs[0]]


def _register_swiss_and_payment_recognizers(engine, languages: list[str]) -> None:
    """Swiss enterprise UID (CHE), AHV/AVS (OASI), SWIFT/BIC.

    CH IBANs are already covered by Presidio's built-in IBAN_CODE (checksum incl. CH).
    Organization *names* still come from NER (ORG / ORGANIZATION); CHE catches the numeric UID.

    One copy of each pattern recognizer per `languages` entry — Presidio only runs recognizers
    whose supported_language matches the analyze() language.
    """
    global _custom_recognizers_registered
    if _custom_recognizers_registered:
        return
    try:
        from presidio_analyzer.pattern import Pattern
        from presidio_analyzer.pattern_recognizer import PatternRecognizer
    except ImportError:
        return

    ch_uid_ctx = [
        "uid",
        "mwst",
        "tva",
        "iva",
        "vat",
        "ide",
        "zefix",
        "handelsregister",
        "handelsregisternummer",
        "company",
        "enterprise",
        "unternehmer",
        "suisse",
        "schweiz",
        "switzerland",
    ]
    ch_ahv_ctx = [
        "ahv",
        "avs",
        "avs/ai",
        "oasi",
        "inps",
        "cassa",
        "assicurazione",
        "ahv-nr",
        "ahvnummer",
        "versicherten",
        "assurance-vieillesse",
        "assurance",
        "sociale",
    ]
    swift_ctx = [
        "swift",
        "bic",
        "iban",
        "bank",
        "banking",
        "clearing",
        "bc-nr",
        "bcnr",
        "zkb",
        "ubs",
        "postfinance",
        "raiffeisen",
        "credit",
        "suisse",
        "schweiz",
        "zahlung",
        "payment",
        "wire",
        "überweisung",
        "virement",
    ]

    for lang in languages:
        ch_uid = PatternRecognizer(
            supported_entity="CH_UID",
            patterns=[
                Pattern("che_uid_dotted", r"\bCHE[-\s]?\d{3}\.\d{3}\.\d{3}\b", 0.92),
                Pattern("che_uid_compact", r"\bCHE[-\s]?\d{9}\b", 0.78),
            ],
            context=ch_uid_ctx,
            supported_language=lang,
        )
        ch_ahv = PatternRecognizer(
            supported_entity="CH_AHV_NUMBER",
            patterns=[
                Pattern("ahv_756_dotted", r"\b756\.\d{4}\.\d{4}\.\d{2}\b", 0.93),
                Pattern("ahv_756_mixed", r"\b756[\s.-]\d{4}[\s.-]\d{4}[\s.-]\d{2}\b", 0.9),
                Pattern("ahv_756_compact", r"\b756\d{10}\b", 0.82),
            ],
            context=ch_ahv_ctx,
            supported_language=lang,
        )
        swift = PatternRecognizer(
            supported_entity="SWIFT_CODE",
            # Slightly higher base score so standalone BIC lines are not dropped when threshold is ~0.2
            patterns=[Pattern("swift_bic", r"\b[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}([A-Z0-9]{3})?\b", 0.5)],
            context=swift_ctx,
            supported_language=lang,
        )
        for rec in (ch_uid, ch_ahv, swift):
            engine.registry.add_recognizer(rec)

    _custom_recognizers_registered = True


def get_engine():
    global _engine, _ENGINE_ANALYSIS_LANGS
    if _engine is None:
        try:
            from presidio_analyzer import AnalyzerEngine
        except ImportError:
            here = os.path.dirname(os.path.abspath(__file__))
            req = os.path.join(here, "requirements-presidio.txt")
            msg = (
                "Presidio is not installed in this Python environment (No module named 'presidio_analyzer'). "
                "From the legal-word-addin folder run: npm run presidio:install "
                f"— or: {sys.executable} -m pip install -r {req} "
                "then: python -m spacy download en_core_web_sm. "
                "If pip fails on macOS, use Python 3.10+ in a venv and set PYTHON in .env to that "
                "interpreter so npm run server spawns the same one. Then restart npm run server."
            )
            print(msg, file=sys.stderr)
            raise ImportError(msg) from None
        nlp_engine, langs = _nlp_engine_and_supported_languages()
        if nlp_engine and langs:
            _engine = AnalyzerEngine(nlp_engine=nlp_engine, supported_languages=langs)
            _ENGINE_ANALYSIS_LANGS = langs
            print(
                f"Presidio: multi-language SpaCy ({', '.join(langs)}). "
                "Install de_core_news_sm / fr_core_news_sm / it_core_news_sm for Swiss DE/FR/IT recall.",
                flush=True,
            )
        else:
            _engine = AnalyzerEngine()
            _ENGINE_ANALYSIS_LANGS = ["en"]
            print(
                "Presidio: English-only SpaCy (en_core_web_sm). "
                "For German/French/Italian documents run: python -m spacy download de_core_news_sm fr_core_news_sm it_core_news_sm",
                flush=True,
            )
        _register_swiss_and_payment_recognizers(_engine, _ENGINE_ANALYSIS_LANGS)
    return _engine


def _merge_recognizer_results(results: list) -> list:
    """Same (start, end) span: keep highest-scoring hit."""
    best: dict[tuple[int, int], object] = {}
    for r in results:
        key = (r.start, r.end)
        prev = best.get(key)
        score = getattr(r, "score", 0) or 0
        if prev is None or score > (getattr(prev, "score", 0) or 0):
            best[key] = r
    return list(best.values())


def analyze_to_items(text: str, score_threshold: float | None = None) -> list[dict]:
    if not text or not text.strip():
        return []
    if score_threshold is None:
        score_threshold = DEFAULT_SCORE_THRESHOLD
    engine = get_engine()
    langs = _filter_analysis_languages(_ENGINE_ANALYSIS_LANGS)
    raw: list = []
    for lang in langs:
        try:
            raw.extend(
                engine.analyze(
                    text=text,
                    language=lang,
                    score_threshold=score_threshold,
                )
            )
        except Exception as ex:
            sys.stderr.write(f"Presidio analyze warning ({lang}): {ex}\n")
    results = _merge_recognizer_results(raw)
    items: list[dict] = []
    seen: set[tuple[str, str]] = set()
    # Longer spans first helps Word-side replacement order match redactItems sort
    for r in sorted(results, key=lambda x: (x.start, -(x.end - x.start))):
        if r.entity_type in SKIP_TYPES:
            continue
        cat = ENTITY_TO_CATEGORY.get(r.entity_type) or DEFAULT_CATEGORY_UNKNOWN_ENTITY
        original = text[r.start : r.end]
        if not original.strip():
            continue
        key = (original.strip().lower(), cat)
        if key in seen:
            continue
        seen.add(key)
        items.append({"original": original, "category": cat})
    # Final sort: longest original first (same as redactItems)
    items.sort(key=lambda i: len(i["original"]), reverse=True)
    return items


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("%s - - [%s] %s\n" % (self.client_address[0], self.log_date_time_string(), fmt % args))

    def _send_json(self, code: int, payload: dict) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if self.path == "/health" or self.path.startswith("/health?"):
            self._send_json(200, {"ok": True})
            return
        self.send_error(404)

    def do_POST(self) -> None:
        if self.path != "/scan":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        raw = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError:
            self._send_json(400, {"error": "invalid_json"})
            return
        text = data.get("text")
        if not isinstance(text, str):
            self._send_json(400, {"error": "missing_text"})
            return
        threshold = data.get("score_threshold", DEFAULT_SCORE_THRESHOLD)
        try:
            t = float(threshold)
        except (TypeError, ValueError):
            t = DEFAULT_SCORE_THRESHOLD
        t = max(0.0, min(1.0, t))
        try:
            items = analyze_to_items(text, score_threshold=t)
        except ImportError as e:
            msg = str(e)
            sys.stderr.write(f"Presidio import error: {msg}\n")
            self._send_json(503, {"error": "presidio_not_installed", "message": msg})
            return
        except Exception as e:
            msg = str(e)
            sys.stderr.write(f"Presidio analyze error: {msg}\n")
            self._send_json(500, {"error": "presidio_failed", "message": msg})
            return
        self._send_json(200, {"items": items})


def main() -> None:
    server = HTTPServer(("127.0.0.1", PORT), Handler)
    print(f"Presidio sidecar on http://127.0.0.1:{PORT} (POST /scan, GET /health)", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
