from modules.ocr import _flatten_points, _iter_from_result


class AmbiguousArray:
    def __init__(self, values):
        self.values = values

    def __bool__(self):
        raise ValueError("ambiguous truth value")

    def __iter__(self):
        return iter(self.values)

    def __len__(self):
        return len(self.values)


class RapidOcrStyleResult:
    def __init__(self):
        self.dt_polys = AmbiguousArray([AmbiguousArray([AmbiguousArray([1, 2]), AmbiguousArray([9, 10])])])
        self.rec_texts = AmbiguousArray(["target"])
        self.rec_scores = AmbiguousArray([0.93])


def test_zero_confidence_is_not_replaced_by_default():
    result = [{"text": "noise", "score": 0.0, "box": [[0, 0], [10, 10]]}]

    parsed = list(_iter_from_result(result))

    assert parsed[0][2] == 0.0


def test_object_result_does_not_truth_test_array_like_values():
    parsed = list(_iter_from_result(RapidOcrStyleResult()))

    points, text, score = parsed[0]
    assert list(points)[0].values == [1, 2]
    assert text == "target"
    assert score == 0.93


def test_flatten_points_accepts_array_like_points():
    points = AmbiguousArray([AmbiguousArray([1.2, 2.8]), AmbiguousArray([9, 10])])

    assert _flatten_points(points) == [(1, 2), (9, 10)]
