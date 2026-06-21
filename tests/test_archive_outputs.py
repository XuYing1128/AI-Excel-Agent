from pathlib import Path

import scripts.archive_old_outputs as archive


def test_archive_old_outputs_moves_without_deleting(tmp_path, monkeypatch):
    outputs = tmp_path / "outputs"
    legacy = outputs / "legacy"
    outputs.mkdir()
    source = outputs / "old.xlsx"
    source.write_bytes(b"example")
    (outputs / "manifest.json").write_text('{"tasks": []}', encoding="utf-8")

    monkeypatch.setattr(archive, "OUTPUTS", outputs)
    monkeypatch.setattr(archive, "LEGACY", legacy)
    moved = archive.archive_old_outputs()

    assert len(moved) == 1
    assert not source.exists()
    assert (legacy / "old.xlsx").read_bytes() == b"example"
    assert (outputs / "manifest.json").exists()
