import re
import string

import pytest

from hwp2pdf.i18n import LANGUAGE_CODES, LANGUAGE_LABELS, TEXT, translate

PLACEHOLDER = re.compile(r"{(\w+)")


def placeholders(text):
    return {name for name in PLACEHOLDER.findall(text)}


def test_language_tables_agree():
    assert set(LANGUAGE_LABELS) == {"ko", "en"}
    assert LANGUAGE_CODES == {label: code for code, label in LANGUAGE_LABELS.items()}


def test_ko_and_en_have_the_same_keys():
    assert set(TEXT["ko"]) == set(TEXT["en"])


@pytest.mark.parametrize("key", sorted(TEXT["ko"]))
def test_placeholders_match_between_languages(key):
    assert placeholders(TEXT["ko"][key]) == placeholders(TEXT["en"][key]), key


@pytest.mark.parametrize("key", sorted(TEXT["ko"]))
def test_every_string_formats_with_its_own_placeholders(key):
    for lang in ("ko", "en"):
        names = placeholders(TEXT[lang][key])
        translate(lang, key, **{name: "x" for name in names})


def test_unknown_language_falls_back_to_korean():
    assert translate("fr", "ready") == TEXT["ko"]["ready"]


def test_unknown_key_returns_the_key():
    assert translate("ko", "no-such-key") == "no-such-key"


def test_no_stray_format_braces():
    # A bare "{" that is not a placeholder would blow up str.format at runtime.
    for lang in ("ko", "en"):
        for key, text in TEXT[lang].items():
            list(string.Formatter().parse(text))
