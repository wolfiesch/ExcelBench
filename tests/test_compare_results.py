from excelbench.harness.runner import compare_results


def test_compare_results_color_case_insensitive():
    expected = {"font_color": "#ff0000"}
    actual = {"font_color": "#FF0000"}
    assert compare_results(expected, actual)


def test_compare_results_missing_expected_none():
    expected = {"border_top": None}
    actual = {}
    assert compare_results(expected, actual)


def test_compare_results_numeric_tolerance():
    expected = {"value": 1.0}
    actual = {"value": 1.00005}
    assert compare_results(expected, actual)

    actual_fail = {"value": 1.01}
    assert not compare_results(expected, actual_fail)
