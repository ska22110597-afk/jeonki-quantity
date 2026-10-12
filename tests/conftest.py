import pytest


@pytest.fixture(autouse=True)
def _isolated_result_dir(request, tmp_path_factory, monkeypatch):
    """테스트가 홈 폴더의 실제 결과 폴더·임시저장 파일을 건드리지 않게 한다."""
    import os

    if request.module.__name__.endswith("test_paths"):
        yield
        return
    if not os.environ.get("JEONKI_RESULT_DIR"):
        monkeypatch.setenv("JEONKI_RESULT_DIR", str(tmp_path_factory.mktemp("result")))
    yield
