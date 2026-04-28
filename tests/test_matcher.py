from modules.matcher import TextQuery, find_text
from modules.models import OCRBox


def boxes():
    return [
        OCRBox("\u5f20\u4e09", 10, 10, 50, 30, 0.95),
        OCRBox("\u786e\u8ba4", 200, 20, 250, 45, 0.98),
        OCRBox("\u786e\u8ba4", 200, 120, 250, 145, 0.97),
        OCRBox("\u5ba1\u6838\u901a\u8fc7", 400, 100, 480, 130, 0.92),
    ]


def test_find_contains_text():
    found = find_text(boxes(), TextQuery(text="\u5f20\u4e09"))
    assert found is not None
    assert found.center_x == 30


def test_find_second_occurrence():
    found = find_text(boxes(), TextQuery(text="\u786e\u8ba4", occurrence=2))
    assert found is not None
    assert found.y1 == 120


def test_find_rightmost():
    found = find_text(boxes(), TextQuery(text="\u5ba1\u6838", position="right"))
    assert found is not None
    assert found.text == "\u5ba1\u6838\u901a\u8fc7"


def test_fuzzy_match():
    found = find_text(boxes(), TextQuery(text="\u5ba1\u67e5\u901a\u8fc7", match="fuzzy", fuzzy_threshold=60))
    assert found is not None
    assert found.text == "\u5ba1\u6838\u901a\u8fc7"


def test_region_match_uses_screen_coordinates():
    cropped_box = OCRBox("\u786e\u8ba4", 10, 10, 50, 30, 0.98, screen_offset_x=100, screen_offset_y=200)
    found = find_text(
        [cropped_box],
        TextQuery(text="\u786e\u8ba4", region={"left": 100, "top": 200, "width": 80, "height": 80}),
    )
    assert found is cropped_box
