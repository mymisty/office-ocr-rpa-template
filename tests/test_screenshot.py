import pytest

from modules.screenshot import normalize_region


def test_normalize_region_from_xyxy():
    assert normalize_region([10, 20, 110, 220]) == {
        "left": 10,
        "top": 20,
        "width": 100,
        "height": 200,
    }


def test_normalize_region_rejects_empty_area():
    with pytest.raises(ValueError):
        normalize_region([10, 20, 10, 220])
