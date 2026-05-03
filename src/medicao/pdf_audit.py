import hashlib
import json
import os
import re
import subprocess
import tempfile
from datetime import datetime

from .utils import normalize


DIGITAL_TEXT_MIN_LEN = 40
CACHE_VERSION = 1
PDF_AUDIT_MAX_PAGES = 80
PDF_AUDIT_MAX_BYTES = 80 * 1024 * 1024
OCR_MAX_PAGES_PER_FILE = 8
_AUDIT_CONTEXT = {
    "analyze_content": False,
    "enable_ocr": False,
    "progress_cb": None,
    "stats": None,
}


def _ocr_enabled():
    if _AUDIT_CONTEXT.get("enable_ocr"):
        return True
    return os.environ.get("MEDICAO_ENABLE_OCR", "").strip().lower() in {"1", "true", "sim", "yes"}


def pdf_content_enabled():
    if _AUDIT_CONTEXT.get("analyze_content"):
        return True
    return os.environ.get("MEDICAO_ANALYZE_PDF_CONTENT", "").strip().lower() in {"1", "true", "sim", "yes"}


def _new_stats():
    return {
        "mode": "profunda com OCR" if _ocr_enabled() else "profunda sem OCR" if pdf_content_enabled() else "rapida por nomes de arquivo",
        "pdf_files": 0,
        "analyzed_files": 0,
        "cache_hits": 0,
        "skipped_large_files": 0,
        "limited_files": 0,
        "pages_seen": 0,
        "pages_processed": 0,
        "digital_pages": 0,
        "ocr_pages": 0,
        "short_text_pages": 0,
        "documents_found": [],
        "errors": [],
    }


def configure_pdf_audit(enable_ocr=False, analyze_content=False, progress_cb=None):
    _AUDIT_CONTEXT["analyze_content"] = bool(analyze_content or enable_ocr)
    _AUDIT_CONTEXT["enable_ocr"] = bool(enable_ocr)
    _AUDIT_CONTEXT["progress_cb"] = progress_cb
    _AUDIT_CONTEXT["stats"] = _new_stats()


def get_pdf_audit_diagnostics():
    return _AUDIT_CONTEXT.get("stats") or _new_stats()


def _stats():
    if _AUDIT_CONTEXT.get("stats") is None:
        _AUDIT_CONTEXT["stats"] = _new_stats()
    return _AUDIT_CONTEXT["stats"]


def _progress(message):
    cb = _AUDIT_CONTEXT.get("progress_cb")
    if cb:
        try:
            cb(message)
        except Exception:
            pass


DOCUMENT_RULES = {
    "CONTRATO": {
        "name": "Contrato de Trabalho",
        "strong": ["contrato de trabalho", "contrato individual de trabalho", "contrato de experiencia"],
        "medium": ["empregador", "empregado", "admissao", "jornada de trabalho", "salario", "clausula"],
        "weak": ["funcao", "assinatura do empregado", "prazo determinado", "prazo indeterminado"],
        "exclude": [],
    },
    "FICHA_REGISTRO": {
        "name": "Ficha de Registro",
        "strong": ["ficha de registro", "registro de empregado", "dados do empregado"],
        "medium": ["ctps", "pis", "data de admissao", "estado civil", "filiacao", "escolaridade"],
        "weak": ["endereco", "dependentes", "funcao"],
        "exclude": [],
    },
    "VT_DECLARACAO": {
        "name": "Termo de Vale Transporte",
        "strong": ["vale transporte", "termo de opcao", "opcao pelo vale transporte"],
        "medium": ["deslocamento residencia trabalho", "renuncia ao vale transporte", "transporte coletivo"],
        "weak": ["desconto de 6", "declaracao de utilizacao", "deslocamento"],
        "exclude": [],
    },
    "EPI": {
        "name": "Termo de Entrega de EPI",
        "strong": ["ficha de entrega de epi", "equipamento de protecao individual"],
        "medium": ["certificado de aprovacao", "treinamento de uso", "recebido em"],
        "weak": ["epi", "devolucao", "substituicao", "assinatura do empregado"],
        "exclude": [],
    },
    "ASO_ADMISSIONAL": {
        "name": "ASO Admissional",
        "strong": ["atestado de saude ocupacional", "exame admissional"],
        "medium": ["aso", "apto", "inapto", "medico do trabalho", "crm", "pcmso"],
        "weak": ["riscos ocupacionais", "admissional"],
        "exclude": [],
    },
    "ASO": {
        "name": "ASO",
        "strong": ["atestado de saude ocupacional"],
        "medium": ["aso", "apto", "inapto", "medico do trabalho", "crm", "pcmso"],
        "weak": ["exame periodico", "exame demissional", "riscos ocupacionais"],
        "exclude": [],
    },
    "FICHA_CADASTRAL": {
        "name": "Ficha Cadastral",
        "strong": ["ficha cadastral", "dados cadastrais"],
        "medium": ["nome completo", "cpf", "rg", "telefone", "email", "banco", "agencia", "conta"],
        "weak": ["endereco"],
        "exclude": ["ficha de registro"],
    },
    "DOCUMENTO_PESSOAL": {
        "name": "Documento Pessoal",
        "strong": ["carteira de identidade", "carteira nacional de habilitacao", "titulo de eleitor"],
        "medium": ["cpf", "rg", "cnh", "comprovante de residencia", "certidao", "reservista"],
        "weak": [],
        "exclude": ["ficha cadastral", "ficha de registro"],
    },
    "TERMO_RESPONSABILIDADE": {
        "name": "Termo de Responsabilidade",
        "strong": ["termo de responsabilidade"],
        "medium": ["declaro estar ciente", "responsabilidade pelo uso", "compromisso", "ciencia", "guarda"],
        "weak": ["zelo", "devolucao"],
        "exclude": [],
    },
    "DECL_GENERICA": {
        "name": "Declaração",
        "strong": ["declaro para os devidos fins", "sob as penas da lei"],
        "medium": ["declaracao", "afirmo", "ciencia"],
        "weak": [],
        "exclude": [],
    },
    "RECIBO": {
        "name": "Recibo",
        "strong": ["recibo"],
        "medium": ["recebi", "valor de", "pagamento", "referente a"],
        "weak": ["assinatura"],
        "exclude": [],
    },
    "ACORDO_COMPENSACAO": {
        "name": "Acordo de Compensação",
        "strong": ["acordo de compensacao", "banco de horas"],
        "medium": ["compensacao de jornada", "horas extras", "sindicato", "acordo individual"],
        "weak": ["jornada"],
        "exclude": [],
    },
    "TERMO_CONFIDENCIALIDADE": {
        "name": "Termo de Confidencialidade",
        "strong": ["termo de confidencialidade"],
        "medium": ["sigilo", "informacoes confidenciais", "confidencial", "nao divulgacao"],
        "weak": ["dados da empresa"],
        "exclude": [],
    },
    "ORDEM_SERVICO": {
        "name": "Ordem de Serviço",
        "strong": ["ordem de servico", "seguranca do trabalho", "nr 01"],
        "medium": ["riscos ocupacionais", "prevencao de acidentes", "normas de seguranca"],
        "weak": ["obrigacoes do empregado"],
        "exclude": [],
    },
    "COMPROVANTE": {
        "name": "Comprovante",
        "strong": ["comprovante"],
        "medium": ["protocolo", "autenticacao", "codigo de validacao"],
        "weak": [],
        "exclude": [],
    },
}


TAG_ALIASES = {
    "ASO": {"ASO", "ASO_ADMISSIONAL"},
}


OCR_SCRIPT = r"""
param([string]$ImagePath)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
function Await($WinRtTask) {
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod }
    $m = $methods | Select-Object -First 1
    $gm = $m.MakeGenericMethod($WinRtTask.GetType().GetGenericArguments()[0])
    $task = $gm.Invoke($null, @($WinRtTask))
    $task.Wait()
    return $task.Result
}
try {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('pt-BR')) }
    if ($null -eq $engine) { exit 1 }
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync([System.IO.Path]::GetFullPath($ImagePath)))
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read))
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream))
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync())
    $result = Await ($engine.RecognizeAsync($bitmap))
    Write-Output $result.Text
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
"""


def _cache_path():
    try:
        import src.config as config
        base = config.APP_DATA_DIR
    except Exception:
        base = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "_dados_app")
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "medicao_pdf_cache.json")


def _file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_cache():
    path = _cache_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        if payload.get("version") == CACHE_VERSION:
            return payload.get("items", {})
    except Exception:
        pass
    return {}


def _save_cache(items):
    try:
        with open(_cache_path(), "w", encoding="utf-8") as f:
            json.dump({"version": CACHE_VERSION, "items": items}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def normalize_document_text(text):
    txt = normalize(text)
    replacements = {
        "trabalh0": "trabalho",
        "empregad0": "empregado",
        "transporle": "transporte",
        "ocupacionai": "ocupacional",
        "ocupaciona1": "ocupacional",
        "ct p s": "ctps",
        "c t p s": "ctps",
        " n r ": " nr ",
    }
    for old, new in replacements.items():
        txt = txt.replace(old, new)
    txt = re.sub(r"[^a-z0-9% .:/-]+", " ", txt)
    return " ".join(txt.split())


def _snippet(text, term):
    pos = text.find(term)
    if pos < 0:
        return term
    start = max(0, pos - 45)
    end = min(len(text), pos + len(term) + 45)
    return text[start:end].strip()


def classify_text(text):
    norm = normalize_document_text(text)
    possible = []
    for tag, rule in DOCUMENT_RULES.items():
        if any(ex in norm for ex in rule.get("exclude", [])):
            continue
        score = 0
        matches = []
        snippets = []
        for weight, bucket in ((7, "strong"), (3, "medium"), (1, "weak")):
            for term in rule.get(bucket, []):
                term_norm = normalize_document_text(term)
                if term_norm and term_norm in norm:
                    score += weight
                    matches.append(term)
                    snippets.append(_snippet(norm, term_norm))

        if score < 5 or not matches:
            continue

        has_strong = any(normalize_document_text(t) in norm for t in rule.get("strong", []))
        if score >= 11 and has_strong:
            confidence = "alta"
        elif score >= 8:
            confidence = "media"
        else:
            confidence = "baixa"

        possible.append({
            "documentType": tag,
            "documentName": rule["name"],
            "score": score,
            "confidence": confidence,
            "matchedKeywords": matches[:8],
            "evidenceSnippets": snippets[:4],
        })

    possible.sort(key=lambda item: item["score"], reverse=True)
    if not possible:
        return {
            "predictedDocumentType": None,
            "confidence": "baixa",
            "matchedKeywords": [],
            "evidenceSnippets": [],
            "possibleTypes": [],
            "warning": "sem evidencias suficientes",
        }

    best = possible[0]
    warning = ""
    if len(possible) > 1 and possible[1]["score"] >= best["score"] - 2:
        warning = "classificacao concorrente proxima"
        if best["confidence"] == "alta":
            best = {**best, "confidence": "media"}

    return {
        "predictedDocumentType": best["documentType"],
        "confidence": best["confidence"],
        "matchedKeywords": best["matchedKeywords"],
        "evidenceSnippets": best["evidenceSnippets"],
        "possibleTypes": possible[:3],
        "warning": warning,
    }


def _ocr_image(img_path):
    try:
        flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", OCR_SCRIPT, "-ImagePath", img_path],
            capture_output=True,
            text=True,
            timeout=90,
            creationflags=flags,
        )
        if result.returncode == 0:
            return result.stdout.strip(), ""
        return "", result.stderr.strip() or "OCR falhou"
    except Exception as exc:
        return "", str(exc)


def _extract_page_text(page, allow_ocr=False):
    errors = []
    digital_text = ""
    try:
        digital_text = page.get_text().strip()
    except Exception as exc:
        errors.append(str(exc))

    if len(digital_text) >= DIGITAL_TEXT_MIN_LEN:
        return {
            "hasDigitalText": True,
            "extractedText": digital_text,
            "ocrText": "",
            "finalText": digital_text,
            "extractionMethod": "texto digital",
            "errors": errors,
        }

    if not allow_ocr:
        return {
            "hasDigitalText": bool(digital_text),
            "extractedText": digital_text,
            "ocrText": "",
            "finalText": digital_text,
            "extractionMethod": "texto curto" if digital_text else "ocr desativado",
            "errors": errors,
        }

    ocr_text = ""
    try:
        import fitz
        pix = page.get_pixmap(dpi=220, colorspace=fitz.csGRAY)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            img_path = tmp.name
        pix.save(img_path)
        try:
            ocr_text, err = _ocr_image(img_path)
            if err:
                errors.append(err)
        finally:
            try:
                os.unlink(img_path)
            except OSError:
                pass
    except Exception as exc:
        errors.append(str(exc))

    final = ocr_text or digital_text
    return {
        "hasDigitalText": bool(digital_text),
        "extractedText": digital_text,
        "ocrText": ocr_text,
        "finalText": final,
        "extractionMethod": "ocr" if ocr_text else "texto curto",
        "errors": errors,
    }


def _group_pages(pages):
    groups = []
    current = None
    for page in pages:
        cls = page.get("classification", {})
        tag = cls.get("predictedDocumentType")
        confidence = cls.get("confidence", "baixa")
        page_no = int(page["pageNumber"])

        if not tag or confidence == "baixa":
            if current and len(page.get("finalText", "")) < 180:
                current["pages"].append(page_no)
            continue

        if current and current["documentType"] == tag and page_no == current["pages"][-1] + 1:
            current["pages"].append(page_no)
            current["score"] += max(1, int(cls.get("possibleTypes", [{}])[0].get("score", 1)))
            current["matchedKeywords"].extend(cls.get("matchedKeywords", []))
            current["evidenceSnippets"].extend(cls.get("evidenceSnippets", []))
        else:
            current = {
                "documentType": tag,
                "documentName": DOCUMENT_RULES.get(tag, {}).get("name", tag),
                "pages": [page_no],
                "confidence": confidence,
                "score": max(1, int(cls.get("possibleTypes", [{}])[0].get("score", 1))),
                "matchedKeywords": list(cls.get("matchedKeywords", [])),
                "evidenceSnippets": list(cls.get("evidenceSnippets", [])),
                "method": page.get("extractionMethod", ""),
            }
            groups.append(current)

    for group in groups:
        group["matchedKeywords"] = list(dict.fromkeys(group["matchedKeywords"]))[:10]
        group["evidenceSnippets"] = list(dict.fromkeys(group["evidenceSnippets"]))[:5]
        if group["score"] >= 14 and group["confidence"] != "baixa":
            group["confidence"] = "alta"
        elif group["score"] >= 8:
            group["confidence"] = "media"
    return groups


def analyze_pdf_file(pdf_path, use_cache=True):
    pdf_path = os.path.abspath(pdf_path)
    stats = _stats()
    stats["pdf_files"] += 1
    cache = _load_cache() if use_cache else {}
    try:
        file_size = os.path.getsize(pdf_path)
        if file_size > PDF_AUDIT_MAX_BYTES:
            stats["skipped_large_files"] += 1
            stats["errors"].append(f"{os.path.basename(pdf_path)} ignorado por tamanho")
            return {
                "filePath": pdf_path,
                "fileName": os.path.basename(pdf_path),
                "totalPages": 0,
                "pages": [],
                "documents": [],
                "errors": [f"PDF ignorado na analise interna por tamanho ({file_size} bytes)"],
            }
        file_hash = _file_hash(pdf_path)
    except OSError as exc:
        return {"filePath": pdf_path, "fileName": os.path.basename(pdf_path), "totalPages": 0, "pages": [], "documents": [], "errors": [str(exc)]}

    cached = cache.get(pdf_path)
    mode = "ocr" if _ocr_enabled() else "safe"
    if cached and cached.get("hash") == file_hash and cached.get("mode") == mode:
        stats["cache_hits"] += 1
        _progress(f"Reutilizando cache: {os.path.basename(pdf_path)}")
        cached_result = cached["result"]
        for doc_item in cached_result.get("documents", []):
            stats["documents_found"].append({
                "fileName": cached_result.get("fileName", os.path.basename(pdf_path)),
                "documentName": doc_item.get("documentName"),
                "documentType": doc_item.get("documentType"),
                "pages": doc_item.get("pages", []),
                "confidence": doc_item.get("confidence"),
                "method": doc_item.get("method"),
                "matchedKeywords": doc_item.get("matchedKeywords", [])[:5],
            })
        return cached["result"]

    result = {
        "filePath": pdf_path,
        "fileName": os.path.basename(pdf_path),
        "hash": file_hash,
        "analyzedAt": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "totalPages": 0,
        "pages": [],
        "documents": [],
        "errors": [],
    }

    doc = None
    try:
        import fitz
        stats["analyzed_files"] += 1
        _progress(f"Abrindo PDF: {os.path.basename(pdf_path)}")
        doc = fitz.open(pdf_path)
        result["totalPages"] = len(doc)
        stats["pages_seen"] += len(doc)
        if len(doc) > PDF_AUDIT_MAX_PAGES:
            stats["limited_files"] += 1
            result["errors"].append(
                f"PDF com {len(doc)} paginas; analise interna limitada as primeiras {PDF_AUDIT_MAX_PAGES}"
            )
        ocr_pages_used = 0
        for idx, page in enumerate(doc, start=1):
            if idx > PDF_AUDIT_MAX_PAGES:
                break
            _progress(f"Processando {os.path.basename(pdf_path)} pagina {idx}/{len(doc)}")
            allow_ocr = _ocr_enabled() and ocr_pages_used < OCR_MAX_PAGES_PER_FILE
            page_data = _extract_page_text(page, allow_ocr=allow_ocr)
            stats["pages_processed"] += 1
            method = page_data.get("extractionMethod")
            if method == "texto digital":
                stats["digital_pages"] += 1
            elif method == "ocr":
                stats["ocr_pages"] += 1
            else:
                stats["short_text_pages"] += 1
            if page_data.get("extractionMethod") == "ocr":
                ocr_pages_used += 1
            final_text = page_data.get("finalText", "")
            classification = classify_text(final_text)
            result["pages"].append({
                "pageNumber": idx,
                "hasDigitalText": page_data.get("hasDigitalText", False),
                "extractedText": page_data.get("extractedText", ""),
                "ocrText": page_data.get("ocrText", ""),
                "finalText": final_text,
                "textLength": len(final_text),
                "extractionMethod": page_data.get("extractionMethod", ""),
                "errors": page_data.get("errors", []),
                "classification": classification,
            })
    except Exception as exc:
        result["errors"].append(str(exc))
        stats["errors"].append(f"{os.path.basename(pdf_path)}: {exc}")
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass

    result["documents"] = _group_pages(result["pages"])
    for doc_item in result["documents"]:
        stats["documents_found"].append({
            "fileName": result["fileName"],
            "documentName": doc_item.get("documentName"),
            "documentType": doc_item.get("documentType"),
            "pages": doc_item.get("pages", []),
            "confidence": doc_item.get("confidence"),
            "method": doc_item.get("method"),
            "matchedKeywords": doc_item.get("matchedKeywords", [])[:5],
        })
    if use_cache:
        cache[pdf_path] = {"hash": file_hash, "mode": mode, "result": result}
        _save_cache(cache)
    return result


def evidence_note(evidence):
    if not evidence:
        return ""
    pages = evidence.get("pages", [])
    if pages:
        page_txt = f"páginas {min(pages)}-{max(pages)}"
    else:
        page_txt = "páginas ?"
    method = evidence.get("method") or "conteúdo"
    confidence = evidence.get("confidence") or "media"
    file_name = evidence.get("fileName") or ""
    keys = ", ".join(evidence.get("matchedKeywords", [])[:4])
    suffix = f" | evidências: {keys}" if keys else ""
    return f"Encontrado em {file_name}, {page_txt}, via {method}, confiança {confidence}{suffix}"


def tags_from_pdf_analysis(analysis, min_confidence="media"):
    tags = set()
    evidence = {}
    allowed = {"alta"} if min_confidence == "alta" else {"alta", "media"}
    for doc in analysis.get("documents", []):
        if doc.get("confidence") not in allowed:
            continue
        tag = doc.get("documentType")
        if not tag:
            continue
        tags.add(tag)
        aliases = TAG_ALIASES.get(tag, set())
        tags.update(aliases)
        item = {
            **doc,
            "filePath": analysis.get("filePath"),
            "fileName": analysis.get("fileName"),
            "totalPages": analysis.get("totalPages"),
            "source": "conteúdo do PDF",
        }
        evidence.setdefault(tag, []).append(item)
        for alias in aliases:
            evidence.setdefault(alias, []).append(item)
    return tags, evidence
