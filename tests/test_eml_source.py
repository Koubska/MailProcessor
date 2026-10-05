from pathlib import Path

from mailprocessor.sources.eml_folder_source import iter_eml_messages, parse_eml_file


def test_iter_eml_messages_reads_single_file(tmp_path: Path) -> None:
    eml_file = tmp_path / "one.eml"
    eml_file.write_text(
        "\n".join(
            [
                "Message-ID: <abc123@example.com>",
                "From: Max Mustermann <max.mustermann@mail.com>",
                "Subject: Schnuppernachmittag",
                "Date: Mon, 23 Nov 2026 14:00:00 +0100",
                "Content-Type: text/plain; charset=utf-8",
                "",
                "Telefonnummer: 1234 567890",
            ]
        ),
        encoding="utf-8",
    )

    messages = list(iter_eml_messages(tmp_path))

    assert len(messages) == 1
    message = messages[0]
    assert message.source_type == "eml"
    assert message.source_location == str(tmp_path.resolve())
    assert message.message_identity == "<abc123@example.com>"
    assert "Telefonnummer: 1234 567890" in message.body_text


def test_iter_eml_messages_uses_fallback_identity_without_message_id(tmp_path: Path) -> None:
    eml_file = tmp_path / "no-id.eml"
    eml_file.write_text(
        "\n".join(
            [
                "From: Max Mustermann <max.mustermann@mail.com>",
                "Subject: Ohne Message ID",
                "Date: Mon, 23 Nov 2026 14:00:00 +0100",
                "Content-Type: text/plain; charset=utf-8",
                "",
                "Hallo Welt",
            ]
        ),
        encoding="utf-8",
    )

    messages = list(iter_eml_messages(tmp_path))

    assert len(messages) == 1
    assert messages[0].message_identity.startswith("fallback:")


def test_iter_eml_messages_extracts_text_plain_from_multipart(tmp_path: Path) -> None:
    eml_file = tmp_path / "multipart.eml"
    eml_file.write_text(
        "\n".join(
            [
                "Message-ID: <multi@example.com>",
                "From: Max Mustermann <max.mustermann@mail.com>",
                "Subject: Multipart",
                "Date: Mon, 23 Nov 2026 14:00:00 +0100",
                'Content-Type: multipart/alternative; boundary="XYZ"',
                "",
                "--XYZ",
                "Content-Type: text/plain; charset=utf-8",
                "",
                "Dies ist der Plain-Text-Teil.",
                "--XYZ",
                "Content-Type: text/html; charset=utf-8",
                "",
                "<p>Dies ist HTML.</p>",
                "--XYZ--",
            ]
        ),
        encoding="utf-8",
    )

    messages = list(iter_eml_messages(tmp_path))

    assert len(messages) == 1
    assert messages[0].body_text == "Dies ist der Plain-Text-Teil."


def test_html_only_mail_is_converted_to_text(tmp_path: Path) -> None:
    path = tmp_path / "html.eml"
    path.write_text(
        "Message-ID: <html@example.com>\nContent-Type: text/html; charset=utf-8\n\n"
        "<html><style>p{}</style><p>Angebot: Experimente</p><p>Tag: Mo &amp; Di</p></html>",
        encoding="utf-8",
    )

    mail = parse_eml_file(path, source_location=str(tmp_path))

    assert mail.body_text == "Angebot: Experimente\n\nTag: Mo & Di"


def test_unknown_charset_falls_back_to_utf8(tmp_path: Path) -> None:
    path = tmp_path / "charset.eml"
    path.write_bytes(
        "Message-ID: <c@example.com>\nContent-Type: text/plain; charset=x-unknown\n\nGrüße".encode()
    )

    assert parse_eml_file(path, source_location=str(tmp_path)).body_text == "Grüße"
