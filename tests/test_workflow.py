import pytest

from modules.workflow import render_step, render_template, validate_template


def test_render_template_value():
    assert render_template("\u70b9\u51fb {{name}}", {"name": "\u5f20\u4e09"}) == "\u70b9\u51fb \u5f20\u4e09"


def test_render_nested_step():
    step = {"action": "click_text", "text": "{{keyword}}", "then": [{"text": "{{name}}"}]}
    rendered = render_step(step, {"name": "\u5f20\u4e09", "keyword": "\u5f20\u4e09\u6309\u94ae"})
    assert rendered["text"] == "\u5f20\u4e09\u6309\u94ae"
    assert rendered["then"][0]["text"] == "\u5f20\u4e09"


def test_validate_template_rejects_missing_text():
    template = {"data_source": {"file": "data/names.csv"}, "steps": [{"action": "click_text"}]}

    with pytest.raises(ValueError, match="requires text"):
        validate_template(template)


def test_validate_template_accepts_nested_branches():
    template = {
        "data_source": {"file": "data/names.csv"},
        "steps": [
            {
                "action": "if_text_exists",
                "text": "\u786e\u8ba4",
                "then": [{"action": "click_text", "text": "\u786e\u8ba4"}],
                "else": [{"action": "scroll", "amount": -5}],
            }
        ],
    }

    validate_template(template)


def test_validate_template_allows_templated_coordinates():
    template = {
        "data_source": {"file": "data/names.csv"},
        "steps": [{"action": "click_xy", "x": "{{x}}", "y": "{{y}}"}],
    }

    validate_template(template)
