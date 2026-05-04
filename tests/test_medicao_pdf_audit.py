from src.medicao.pdf_audit import classify_text, evidence_note, tags_from_pdf_analysis
from src.medicao.auditor import audit_ferias, audit_inss_fgts, run_audit
from src.medicao.report import generate_report
import src.medicao.auditor as auditor
import src.medicao.pdf_audit as pdf_audit
import src.medicao.pdf_reader as pdf_reader
import re


def test_classify_contract_with_multiple_evidences():
    result = classify_text(
        """
        CONTRATO INDIVIDUAL DE TRABALHO
        Pelo presente instrumento, empregador e empregado ajustam data de admissao,
        funcao, salario, jornada de trabalho e demais clausulas.
        """
    )

    assert result["predictedDocumentType"] == "CONTRATO"
    assert result["confidence"] == "alta"
    assert "contrato individual de trabalho" in result["matchedKeywords"]


def test_filename_variations_for_rescisao_and_ferias_are_accepted(tmp_path):
    rescisao = tmp_path / "3. COMPROVANTE DE RECISÃO.pdf"
    ferias = tmp_path / "Férias Acre 05.02.pdf"
    liq_ferias = tmp_path / "LIQ FERIAS ACRE-02-2026.pdf"
    for path in (rescisao, ferias, liq_ferias):
        path.write_bytes(b"%PDF-1.4")

    assert "COMPROVANTE_RESCISAO" in pdf_reader.identify_doc_types(str(rescisao))
    assert "COMPROVANTE_FERIAS" in pdf_reader.identify_doc_types(str(ferias))
    assert "RECIBO_FERIAS" not in pdf_reader.identify_doc_types(str(ferias))
    assert "RECIBO_FERIAS" not in pdf_reader.identify_doc_types(str(liq_ferias))


def test_ferias_requires_pdf_content_to_confirm_recibo_vs_comprovante():
    recibo = classify_text(
        """
        RECIBO DE FÉRIAS
        Declaro que recebi o valor referente ao gozo de férias e 1/3 constitucional.
        Assinatura do empregado.
        """
    )
    comprovante = classify_text(
        """
        COMPROVANTE DE PAGAMENTO
        Transferência PIX realizada para favorecido. Valor pago referente a férias.
        Autenticação bancária.
        """
    )

    assert recibo["predictedDocumentType"] == "RECIBO_FERIAS"
    assert comprovante["predictedDocumentType"] == "COMPROVANTE_FERIAS"


def test_vt_page_text_is_classified_inside_large_contract():
    result = classify_text(
        """
        TERMO DE OPÇÃO DE VALE TRANSPORTE
        Eu, VINICIUS MOREIRA DE BARROS, declaro para os efeitos do benefício do Vale-Transporte
        que não opto pela utilização do Vale-Transporte.
        Os meios de transporte coletivo, público e regular são adequados.
        """
    )

    assert result["predictedDocumentType"] == "VT_DECLARACAO"


def test_fast_mode_does_not_open_pdf_content(monkeypatch, tmp_path):
    pdf_path = tmp_path / "CONTRATO TRABALHO.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")

    monkeypatch.setattr(pdf_reader, "pdf_content_enabled", lambda: False)

    def fail_if_called(_path):
        raise AssertionError("fast mode must not open PDF internals")

    monkeypatch.setattr(pdf_reader, "analyze_pdf_file", fail_if_called)

    tags = pdf_reader.get_all_tags_in_folder(str(tmp_path))

    assert "CONTRATO" in tags


def test_run_audit_accepts_admissoes_tuple_return(monkeypatch, tmp_path):
    monkeypatch.setattr(auditor, "load_config", lambda: {})
    monkeypatch.setattr(auditor, "detect_competencia", lambda _folder: (3, 2026, "MARÇO"))
    monkeypatch.setattr(auditor, "find_forca_trabalho", lambda _folder: None)
    monkeypatch.setattr(auditor, "find_folder", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(auditor, "list_subfolders", lambda _folder: [])
    monkeypatch.setattr(auditor, "has_any_file", lambda _folder: False)
    monkeypatch.setattr(auditor, "audit_acordo_coletivo", lambda _folder: {"name": "Acordo", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_declaracoes", lambda *_args, **_kwargs: {"name": "Declarações", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_admissoes", lambda *_args, **_kwargs: ({"name": "Admissões", "employees": [], "issues": [], "status": "ok"}, []))
    monkeypatch.setattr(auditor, "audit_demissoes", lambda *_args, **_kwargs: {"name": "Demissões", "employees": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_ferias", lambda *_args, **_kwargs: {"name": "Férias", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_fopag", lambda _folder: {"name": "FOPAG", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_inss_fgts", lambda *_args, **_kwargs: {"name": "INSS", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_ponto", lambda *_args, **_kwargs: {"name": "Ponto", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_va_vr", lambda _folder: {"name": "VA/VR", "items": [], "issues": [], "status": "ok"})
    monkeypatch.setattr(auditor, "audit_vt", lambda _folder: {"name": "VT", "items": [], "issues": [], "status": "ok"})

    result = run_audit(str(tmp_path))

    assert all(isinstance(section, dict) for section in result["sections"])


def test_classify_does_not_accept_single_generic_transport_word():
    result = classify_text("O colaborador utiliza transporte diariamente.")

    assert result["predictedDocumentType"] is None


def test_tags_from_pdf_analysis_keeps_medium_and_high_confidence_only():
    analysis = {
        "filePath": "C:/tmp/documentos.pdf",
        "fileName": "documentos.pdf",
        "totalPages": 4,
        "documents": [
            {
                "documentType": "VT_DECLARACAO",
                "documentName": "Termo de Vale Transporte",
                "pages": [2, 3],
                "confidence": "media",
                "method": "ocr",
                "matchedKeywords": ["vale transporte", "deslocamento"],
                "evidenceSnippets": [],
            },
            {
                "documentType": "RECIBO",
                "documentName": "Recibo",
                "pages": [4],
                "confidence": "baixa",
                "method": "texto digital",
                "matchedKeywords": ["recibo"],
                "evidenceSnippets": [],
            },
        ],
    }

    tags, evidence = tags_from_pdf_analysis(analysis)

    assert tags == {"VT_DECLARACAO"}
    assert evidence["VT_DECLARACAO"][0]["pages"] == [2, 3]
    note = evidence_note(evidence["VT_DECLARACAO"][0])
    assert "documentos.pdf" in note
    assert "C:/tmp/documentos.pdf" in note


def test_ferias_summary_keeps_employee_context_from_subfolder(monkeypatch):
    monkeypatch.setattr(auditor, "find_folder", lambda *_: "FERIAS")
    monkeypatch.setattr(auditor, "list_subfolders", lambda path: ["FERIAS/VINICIUS MOREIRA"] if path == "FERIAS" else [])
    monkeypatch.setattr(auditor, "get_all_tags_in_folder", lambda *_args, **_kwargs: {"AVISO_FERIAS"})

    section = audit_ferias("COMPETENCIA", [])

    assert section["employees"][0]["name"] == "VINICIUS MOREIRA"
    assert "VINICIUS MOREIRA: Recibo de férias não encontrado" in [
        issue["msg"] for issue in section["issues"]
    ]


def test_inss_fgts_summary_marks_company_scope(monkeypatch):
    monkeypatch.setattr(auditor, "find_folder", lambda *_: "INSS + FGTS")
    monkeypatch.setattr(auditor, "get_all_tags_in_folder", lambda *_args, **_kwargs: set())

    section = audit_inss_fgts("COMPETENCIA", {})

    assert any(
        issue["msg"] == "Competência/empresa: Comprovante de pagamento FGTS não encontrado"
        for issue in section["issues"]
    )


def test_report_uses_same_justification_key_for_summary_and_detail(tmp_path):
    audit_result = {
        "competencia": "MARCO/2026",
        "folder_path": str(tmp_path),
        "timestamp": "03/05/2026 10:00",
        "overall_status": "warning",
        "all_issues": [
            {
                "msg": "Demissões — MIZAEL CARVALHO DOS SANTOS: Aviso prévio não encontrado — verificar se se aplica",
                "section": "Demissões",
            }
        ],
        "sections": [
            {
                "name": "Demissões",
                "icon": "🔚",
                "status": "warning",
                "employees": [
                    {
                        "name": "MIZAEL CARVALHO DOS SANTOS",
                        "status": "warning",
                        "issues": ["Aviso prévio não encontrado — verificar se se aplica"],
                        "items": [
                            {
                                "label": "Aviso prévio / Pedido de demissão",
                                "status": "warning",
                                "note": "Não encontrado (verifique tipo de contrato)",
                            }
                        ],
                    }
                ],
                "issues": [],
            }
        ],
    }

    report_path = generate_report(audit_result, str(tmp_path))
    html = open(report_path, encoding="utf-8").read()
    keys = [
        key for key in re.findall(r'data-justifiable="([^"]+)"', html)
        if re.fullmatch(r'[0-9a-f]{10}', key)
    ]

    assert len(keys) == 2
    assert len(set(keys)) == 1


def test_report_deduplicates_issues_and_explains_pending_reason(tmp_path):
    duplicate_msg = "FOPAG — Competência/empresa: Comprovante de pagamento de salário não encontrado"
    audit_result = {
        "competencia": "MARCO/2026",
        "folder_path": str(tmp_path),
        "timestamp": "03/05/2026 10:00",
        "overall_status": "error",
        "all_issues": [
            {"msg": duplicate_msg, "section": "FOPAG"},
            {"msg": duplicate_msg, "section": "FOPAG"},
        ],
        "sections": [],
    }

    report_path = generate_report(audit_result, str(tmp_path))
    html = open(report_path, encoding="utf-8").read()

    assert html.count(duplicate_msg) == 1
    assert "Por que isso é pendência" in html
    assert "Ocorrências agrupadas" in html


def test_report_shows_file_path_from_evidence_note(tmp_path):
    audit_result = {
        "competencia": "MARCO/2026",
        "folder_path": str(tmp_path),
        "timestamp": "03/05/2026 10:00",
        "overall_status": "warning",
        "all_issues": [],
        "sections": [
            {
                "name": "Diagnóstico PDF",
                "status": "warning",
                "items": [
                    {
                        "label": "Ficha Cadastral em documento.pdf",
                        "status": "warning",
                        "note": "páginas 1-2; confiança media; caminho: C:/medicoes/MARCO/documento.pdf",
                    }
                ],
                "issues": [],
            }
        ],
    }

    report_path = generate_report(audit_result, str(tmp_path))
    html = open(report_path, encoding="utf-8").read()

    assert "Endereço do arquivo" in html
    assert "C:/medicoes/MARCO/documento.pdf" in html


def test_deep_search_resolves_pending_document_found_elsewhere(monkeypatch, tmp_path):
    sections = [
        {
            "name": "INSS + FGTS",
            "status": "error",
            "items": [
                {"label": "Comprovante de pagamento FGTS", "status": "error", "note": "FALTANDO"},
            ],
            "issues": [
                {"msg": "Competência/empresa: Comprovante de pagamento FGTS não encontrado", "section": "INSS + FGTS"},
            ],
        }
    ]
    monkeypatch.setattr(
        auditor,
        "get_pdf_evidence_in_folder",
        lambda *_args, **_kwargs: {
            "COMPROVANTE_FGTS": [
                {
                    "filePath": str(tmp_path / "FGTS DETALHAMENTO-MIZAEL CARVALHO.pdf"),
                    "fileName": "FGTS DETALHAMENTO-MIZAEL CARVALHO.pdf",
                    "pages": [4],
                    "confidence": "media",
                    "method": "texto digital",
                    "matchedKeywords": ["fgts", "pagamento"],
                }
            ]
        },
    )

    auditor._resolve_pendencias_por_busca_profunda(sections, str(tmp_path))

    assert sections[0]["items"][0]["status"] == "ok"
    assert "Busca profunda encontrou" in sections[0]["items"][0]["note"]
    assert sections[0]["issues"] == []
    assert sections[0]["status"] == "ok"


def test_deep_search_resolves_vt_by_filename_for_employee(tmp_path, monkeypatch):
    person_dir = tmp_path / "Admissão-alocação" / "VINICIUS MOREIRA DE BARROS"
    person_dir.mkdir(parents=True)
    (person_dir / "VALE TRANSPORTE - VINICIUS.pdf").write_bytes(b"%PDF-1.4")
    sections = [
        {
            "name": "Admissões",
            "status": "error",
            "employees": [
                {
                    "name": "VINICIUS MOREIRA DE BARROS",
                    "status": "error",
                    "items": [{"label": "Declaração de opção VT", "status": "warning", "note": "Não encontrado"}],
                    "issues": ["Declaração de opção VT não encontrada"],
                }
            ],
            "issues": [],
        }
    ]
    monkeypatch.setattr(auditor, "get_pdf_evidence_in_folder", lambda *_args, **_kwargs: {})

    auditor._resolve_pendencias_por_busca_profunda(sections, str(tmp_path))

    item = sections[0]["employees"][0]["items"][0]
    assert item["status"] == "ok"
    assert "VALE TRANSPORTE - VINICIUS.pdf" in item["note"]
    assert sections[0]["employees"][0]["issues"] == []


def test_force_reprocess_allows_ocr_on_pages_beyond_default_limit():
    pdf_audit.configure_pdf_audit(enable_ocr=True, force_reprocess=False)
    limited = pdf_audit.get_pdf_audit_diagnostics()
    pdf_audit.configure_pdf_audit(enable_ocr=True, force_reprocess=True)
    full = pdf_audit.get_pdf_audit_diagnostics()

    assert "8 paginas" in limited["ocr_scope"]
    assert full["ocr_scope"] == "todas as paginas analisadas"


def test_report_shows_ok_when_only_technical_warnings_exist(tmp_path):
    audit_result = {
        "competencia": "MARCO/2026",
        "folder_path": str(tmp_path),
        "timestamp": "03/05/2026 10:00",
        "overall_status": "warning",
        "all_issues": [],
        "sections": [
            {
                "name": "Diagnóstico da Auditoria PDF",
                "status": "warning",
                "items": [
                    {
                        "label": "Ficha Cadastral em documento.pdf",
                        "status": "info",
                        "note": "confiança media; achado tecnico para conferencia, nao e pendencia",
                    }
                ],
                "issues": [],
            }
        ],
    }

    report_path = generate_report(audit_result, str(tmp_path))
    html = open(report_path, encoding="utf-8").read()

    assert "STATUS GERAL: OK PARA ENVIO" in html
    assert "Nenhuma pendência encontrada" in html


def test_pdf_diagnostic_labels_explain_cache_usage():
    section = auditor._pdf_diagnostic_section(
        {
            "mode": "profunda com OCR",
            "pdf_files": 148,
            "analyzed_files": 0,
            "cache_hits": 148,
            "pages_processed": 0,
            "digital_pages": 0,
            "ocr_pages": 0,
            "documents_found": [],
            "errors": [],
        }
    )

    labels = {item["label"]: item["note"] for item in section["items"]}
    assert labels["PDFs considerados na auditoria"] == "148"
    assert labels["PDFs reprocessados nesta execucao"] == "0"
    assert labels["PDFs reutilizados do cache"] == "148"


def test_pdf_diagnostic_shows_when_cache_is_ignored():
    pdf_audit.configure_pdf_audit(force_reprocess=True)
    section = auditor._pdf_diagnostic_section(pdf_audit.get_pdf_audit_diagnostics())
    labels = {item["label"]: item["note"] for item in section["items"]}

    assert labels["Cache da medicao"] == "ignorado nesta execucao"
