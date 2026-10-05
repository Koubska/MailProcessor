"""Guards for the GitHub Actions workflow (builds and releases only for tags, after the tests)."""

import re
from pathlib import Path

WORKFLOW = Path(__file__).resolve().parents[1] / ".github" / "workflows" / "ci.yml"
TAG_CONDITION = "if: startsWith(github.ref, 'refs/tags/')"


def _job(name: str) -> str:
    """Text of one job: from `  name:` up to the next job at the same indentation."""
    content = WORKFLOW.read_text(encoding="utf-8")
    start = content.index(f"\n  {name}:\n")
    following = [content.find(f"\n  {other}:\n", start + 1) for other in ("test", "build", "release")]
    end = min((pos for pos in following if pos > start), default=len(content))
    return content[start:end]


def test_tests_run_on_pushes_tags_and_pull_requests_on_all_platforms() -> None:
    content = WORKFLOW.read_text(encoding="utf-8")
    test_job = _job("test")

    assert 'tags: ["*"]' in content
    assert "pull_request:" in content
    assert "os: [ubuntu-latest, windows-latest, macos-latest]" in test_job
    assert "pytest -q" in test_job
    assert TAG_CONDITION not in test_job


def test_builds_only_for_tags_after_tests_on_every_platform() -> None:
    build_job = _job("build")

    assert TAG_CONDITION in build_job
    assert "needs: test" in build_job
    for os_name, platform in (("windows-latest", "windows"), ("macos-latest", "macos"), ("ubuntu-latest", "linux")):
        assert re.search(rf"- os: {os_name}( +#[^\n]*)?\n +platform: {platform}\n", build_job)
    assert "python scripts/build_executable.py --target-platform ${{ matrix.platform }}" in build_job
    assert "path: dist/mailprocessor-${{ matrix.platform }}.zip" in build_job
    assert "if-no-files-found: error" in build_job


def test_release_publishes_all_zips_only_for_tags() -> None:
    release_job = _job("release")

    assert TAG_CONDITION in release_job
    assert "needs: build" in release_job
    assert "contents: write" in release_job
    assert "merge-multiple: true" in release_job
    assert 'gh release create "$GITHUB_REF_NAME" dist/mailprocessor-*.zip' in release_job


def test_workflow_has_read_only_default_permissions() -> None:
    content = WORKFLOW.read_text(encoding="utf-8")

    assert "permissions:\n  contents: read\n" in content
