from src.medicao.pdf_audit import classify_text, evidence_note, tags_from_pdf_analysis
from src.medicao.auditor import audit_ferias, audit_inss_fgts, run_audit
from src.medicao.report import generate_report
import src.medicao.auditor as auditor
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
    assert "documentos.pdf" in evidence_note(evidence["VT_DECLARACAO"][0])


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
