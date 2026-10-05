"""Guards for the GitHub Actions workflow (builds and releases only for tags, after the tests)."""

import re
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
WORKFLOW = REPO / ".github" / "workflows" / "ci.yml"
TAG_CONDITION = "if: startsWith(github.ref, 'refs/tags/')"


def _job(name: str) -> str:
    """Text of one job: from `  name:` up to the next job at the same indentation."""
    content = WORKFLOW.read_text(encoding="utf-8")
    start = content.index(f"\n  {name}:\n")
    jobs = ("test", "lint", "build", "release", "dependabot-merge")
    following = [content.find(f"\n  {other}:\n", start + 1) for other in jobs]
    end = min((pos for pos in following if pos > start), default=len(content))
    return content[start:end]


def test_tests_run_on_pushes_tags_and_pull_requests_on_all_platforms() -> None:
    content = WORKFLOW.read_text(encoding="utf-8")
    test_job = _job("test")

    assert 'tags: ["*"]' in content
    assert "pull_request:" in content
    assert "os: [ubuntu-latest, windows-latest, macos-latest]" in test_job
    assert "uv sync --locked" in test_job
    assert "uv run --locked pytest -q" in test_job
    assert TAG_CONDITION not in test_job


def test_lint_checks_code_and_workflows() -> None:
    lint_job = _job("lint")

    assert "uv run --locked ruff check src tests scripts" in lint_job
    assert "actionlint" in lint_job
    assert TAG_CONDITION not in lint_job


def test_builds_only_for_tags_after_tests_on_every_platform() -> None:
    build_job = _job("build")

    assert "startsWith(github.ref, 'refs/tags/')" in build_job
    assert "needs: [test, lint]" in build_job
    # Only releases are attested; Dependabot and "build"-labelled PRs are built without it.
    step = "actions/attest-build-provenance"
    step_text = build_job[build_job.rindex("- ", 0, build_job.index(step)) : build_job.index(step)]
    assert TAG_CONDITION in step_text
    # The version step runs in every build (so it is tested before a release), with the tag on releases.
    start, end = build_job.index("- name: Set the version"), build_job.index("scripts/build_executable.py")
    version_step = build_job[start:end]
    assert "if:" not in version_step
    assert "startsWith(github.ref, 'refs/tags/') && github.ref_name" in version_step
    assert 'uv version "${VERSION#v}" --frozen' in version_step
    # bash syntax like ${VERSION#v} needs bash on Windows too (the default there is PowerShell).
    assert re.search(r"\n    defaults:\n      run:\n(?: +#[^\n]*\n)* +shell: bash\n", build_job)
    assert "contains(github.event.pull_request.labels.*.name, 'build')" in build_job
    assert "subject-path: dist/mailprocessor-${{ matrix.platform }}.zip" in build_job
    assert "id-token: write" in build_job and "attestations: write" in build_job
    assert "uv sync --locked --group build" in build_job
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
    assert "--generate-notes" in release_job


def test_workflow_has_read_only_default_permissions() -> None:
    content = WORKFLOW.read_text(encoding="utf-8")

    assert "permissions:\n  contents: read\n" in content


def test_dependabot_updates_python_packages_and_actions() -> None:
    config = (REPO / ".github" / "dependabot.yml").read_text(encoding="utf-8")

    assert 'package-ecosystem: "uv"' in config
    assert 'package-ecosystem: "github-actions"' in config
    assert (REPO / "uv.lock").is_file()


def test_dependabot_merges_only_small_updates_after_everything_passed() -> None:
    merge_job = _job("dependabot-merge")

    assert "needs: [test, lint, build]" in merge_job
    assert "github.event.pull_request.user.login == 'dependabot[bot]'" in merge_job
    assert "version-update:semver-patch" in merge_job and "version-update:semver-minor" in merge_job
    assert "semver-major" not in merge_job
    assert TAG_CONDITION not in merge_job
    # Release stays tag-only.
    assert _job("release").count("dependabot") == 0
