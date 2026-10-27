"""Live knowledge ingestion tests (FR7): Tavily + Grok fill a KB miss, get embedded,
retrieved and verified; and the system is honest (never cites a wrong-year doc) when no
live source is configured."""
import freshguard.live_sources as ls
from freshguard.pipeline import FreshGuard
from freshguard.storage import Storage


class _Resp:
    def __init__(self, payload, status=200):
        self._p, self.status_code = payload, status

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"http {self.status_code}")

    def json(self):
        return self._p


def _fake_post(url, json=None, headers=None, timeout=None):
    if "tavily" in url:
        return _Resp({
            "answer": "France won the 2026 FIFA World Cup, defeating Argentina in 2026.",
            "results": [{
                "title": "2026 World Cup final",
                "url": "https://news.example/wc2026",
                "content": "In the 2026 FIFA World Cup final, France beat Argentina to win the title.",
                "published_date": "2026-07-19",
            }],
        })
    if "x.ai" in url:
        return _Resp({"choices": [{"message": {
            "content": "France won the 2026 FIFA World Cup in 2026, defeating Argentina in the final."}}]})
    return _Resp({}, 404)


# -- offline: KB miss on a future year must NOT answer from a stale same-topic doc ----
def test_future_year_without_keys_is_honest(fg, monkeypatch):
    monkeypatch.delenv("TAVILY_API_KEY", raising=False)
    monkeypatch.delenv("XAI_API_KEY", raising=False)
    monkeypatch.delenv("GROK_API_KEY", raising=False)
    r = fg.answer("Who won the 2026 FIFA World Cup?", persist=False)
    # it must not cite the 2022 document, and must not claim Argentina/France won 2026
    assert all(s["doc_id"] != "d001" for s in r["sources"])
    assert "2022" not in r["answer"]
    assert r["knowledge_sources"] == []
    assert any("not in the knowledge base" in w.lower() or "live source" in w.lower()
               for w in r["warnings"] + [r["answer"]])


# -- live: Tavily + Grok fill the KB, then retrieve + verify --------------------------
def test_live_ingestion_fills_kb_and_verifies(monkeypatch, tmp_path):
    monkeypatch.setenv("TAVILY_API_KEY", "test")
    monkeypatch.setenv("XAI_API_KEY", "test")
    monkeypatch.setattr(ls.requests, "post", _fake_post)

    storage = Storage(tmp_path / "live.db")
    fg = FreshGuard(storage=storage)
    r = fg.answer("Who won the 2026 FIFA World Cup?", persist=False)

    assert set(r["knowledge_sources"]) == {"tavily", "grok"}   # both sources used
    assert r["live_ingested"], "expected docs ingested on a KB miss"
    assert "france" in r["answer"].lower() and "2026" in r["answer"]
    assert all(s["doc_id"] != "d001" for s in r["sources"])    # stale doc excluded
    assert any(s["label"] == "supported" for s in r["statements"])
    assert r["confidence"] > 60
    # the fetched knowledge was persisted into the KB (embedded for future retrieval)
    assert any(d["doc_id"] in r["live_ingested"] for d in storage.kb_all())


# -- sources are inert without keys (keeps the suite offline/deterministic) -----------
def test_sources_unavailable_without_keys(monkeypatch):
    monkeypatch.delenv("TAVILY_API_KEY", raising=False)
    monkeypatch.delenv("XAI_API_KEY", raising=False)
    monkeypatch.delenv("GROK_API_KEY", raising=False)
    assert ls.tavily_available() is False
    assert ls.grok_available() is False
    assert ls.tavily_search("anything") == []
    assert ls.grok_knowledge("anything") == []
