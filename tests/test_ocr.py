from modules.ocr import _iter_from_result


def test_zero_confidence_is_not_replaced_by_default():
    result = [{"text": "noise", "score": 0.0, "box": [[0, 0], [10, 10]]}]

    parsed = list(_iter_from_result(result))

    assert parsed[0][2] == 0.0
