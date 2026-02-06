from excelbench.harness.runner import calculate_score
from excelbench.models import Importance, OperationType, TestResult


def _make_result(passed: bool, importance: Importance) -> TestResult:
    return TestResult(
        test_case_id="case",
        operation=OperationType.READ,
        passed=passed,
        expected={},
        actual={},
        importance=importance,
    )


def test_score_all_basic_and_edge_pass() -> None:
    results = [
        _make_result(True, Importance.BASIC),
        _make_result(True, Importance.EDGE),
    ]
    assert calculate_score(results) == 3


def test_score_basic_pass_edge_fail() -> None:
    results = [
        _make_result(True, Importance.BASIC),
        _make_result(False, Importance.EDGE),
    ]
    assert calculate_score(results) == 2


def test_score_partial_basic() -> None:
    results = [
        _make_result(True, Importance.BASIC),
        _make_result(False, Importance.BASIC),
        _make_result(True, Importance.EDGE),
    ]
    assert calculate_score(results) == 1


def test_score_no_basic_pass() -> None:
    results = [
        _make_result(False, Importance.BASIC),
        _make_result(True, Importance.EDGE),
    ]
    assert calculate_score(results) == 0
