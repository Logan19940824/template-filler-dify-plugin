from pathlib import Path

import pytest

from template_filler.services.excel import parse_excel
from template_filler.services.word import parse_word


@pytest.mark.parametrize("parser", [parse_word, parse_excel])
def test_rejects_wrong_file_content(tmp_path: Path, parser):
    path = tmp_path / "preview"
    path.write_text("<html><body>preview</body></html>", encoding="utf-8")
    with pytest.raises(ValueError, match="valid (Word DOCX|Excel XLSX)"):
        parser(path)
