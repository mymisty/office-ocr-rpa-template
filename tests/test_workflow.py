from modules.workflow import render_step, render_template


def test_render_template_value():
    assert render_template("\u70b9\u51fb {{name}}", {"name": "\u5f20\u4e09"}) == "\u70b9\u51fb \u5f20\u4e09"


def test_render_nested_step():
    step = {"action": "click_text", "text": "{{keyword}}", "then": [{"text": "{{name}}"}]}
    rendered = render_step(step, {"name": "\u5f20\u4e09", "keyword": "\u5f20\u4e09\u6309\u94ae"})
    assert rendered["text"] == "\u5f20\u4e09\u6309\u94ae"
    assert rendered["then"][0]["text"] == "\u5f20\u4e09"
