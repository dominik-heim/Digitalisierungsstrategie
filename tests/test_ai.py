"""KI-Modul ohne echten API-Aufruf testen (Client wird ersetzt)."""
import json
from types import SimpleNamespace

import pytest

from digistrat import ai
from digistrat.model import METHODS


class FakeStream:
    def __init__(self, payload, captured, **kw):
        self.payload = payload
        captured.update(kw)

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_final_message(self):
        return SimpleNamespace(stop_reason="end_turn", content=[SimpleNamespace(type="text", text=json.dumps(self.payload))])


def fake_client(payload, captured):
    return SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(
        stream=lambda **kw: FakeStream(payload, captured, **kw))))


def test_schema_is_strict():
    for key, m in METHODS.items():
        if m.wide:
            continue
        s = ai.rows_schema(m)
        item = s["properties"]["rows"]["items"]
        assert item["additionalProperties"] is False
        assert set(item["required"]) == set(item["properties"])


def test_extract_rows(monkeypatch, tmp_path):
    captured = {}
    payload = {"rows": [{f.key: None if f.type in ("int", "float") else "" for f in METHODS["painpoints"].fields}
                        | {"id": "PP-01", "beschreibung": "Doppelerfassung"}], "hinweise": ""}
    monkeypatch.setattr(ai, "_client", lambda: fake_client(payload, captured))
    note = tmp_path / "notiz.txt"
    note.write_text("Rote Stickies: Doppelerfassung Excel/ERP")
    rows, _ = ai.extract_rows("painpoints", "weitere Notizen", [note], {"ki": {"model": "claude-opus-5-5"}})
    assert rows[0]["id"] == "PP-01"
    assert captured["model"] == "claude-opus-5-5"
    assert captured["output_config"]["format"]["type"] == "json_schema"
    assert captured["fallbacks"] == "default"
    texts = [b["text"] for b in captured["messages"][0]["content"] if b["type"] == "text"]
    assert any("Doppelerfassung Excel/ERP" in t for t in texts)


def test_refusal(monkeypatch):
    class R(FakeStream):
        def get_final_message(self):
            return SimpleNamespace(stop_reason="refusal", content=[])
    monkeypatch.setattr(ai, "_client", lambda: SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(
        stream=lambda **kw: R({}, {}, **kw)))))
    with pytest.raises(ai.AIError):
        ai.ask_json("x", {"type": "object"})


def test_ki_deaktiviert_sendet_nichts(monkeypatch):
    def boom():
        raise AssertionError("Client darf nicht erzeugt werden")
    monkeypatch.setattr(ai, "_client", boom)
    cfg = {"ki": {"aktiv": False}}
    with pytest.raises(ai.AIError):
        ai.extract_rows("painpoints", "x", [], cfg)
    with pytest.raises(ai.AIError):
        ai.derive_rows("userstories", {}, cfg)
    with pytest.raises(ai.AIError):
        ai.ai_management_summary({}, cfg)
    with pytest.raises(ai.AIError):
        ai.workshop_questions(cfg, [])


def test_describe_payload_markiert_personendaten():
    res = {"stakeholder": {"items": [1, 2]}, "painpoints": {"items": [1]}, "issues": []}
    lines = ai.describe_payload("summary", {"ki": {"model": "claude-opus-5-5"}}, results=res)
    assert any(x.startswith("⚠ Stakeholder") for x in lines)
    assert any("Pain Points (1)" in x for x in lines)
    assert lines[-1].startswith("Empfänger: Anthropic")
    assert not any("Stakeholder" in x for x in ai.describe_payload("derive", {}, results=res))
