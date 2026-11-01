"""규격 파서와 판정 로직 단위 테스트.  실행: python -m unittest -v"""

import unittest

from db import guess_judge_type, judge_values, parse_spec


class TestParseSpec(unittest.TestCase):
    def test_symmetric(self):
        self.assertEqual(parse_spec("1.25±0.10"), (1.15, 1.35))
        self.assertEqual(parse_spec("1.25 ± 0.1"), (1.15, 1.35))

    def test_asymmetric(self):
        self.assertEqual(parse_spec("1.25+0.10/-0.05"), (1.20, 1.35))
        self.assertEqual(parse_spec("10 +0.2 -0.3"), (9.7, 10.2))

    def test_range(self):
        self.assertEqual(parse_spec("1.0~2.0"), (1.0, 2.0))
        self.assertEqual(parse_spec("2.0 ～ 1.0"), (1.0, 2.0))

    def test_upper_only(self):
        self.assertEqual(parse_spec("MAX 2.0"), (None, 2.0))
        self.assertEqual(parse_spec("2.0 max"), (None, 2.0))
        self.assertEqual(parse_spec("0.5 이하"), (None, 0.5))
        self.assertEqual(parse_spec("≤ 3.5"), (None, 3.5))

    def test_lower_only(self):
        self.assertEqual(parse_spec("MIN 1.0"), (1.0, None))
        self.assertEqual(parse_spec("3.0 이상"), (3.0, None))
        self.assertEqual(parse_spec("≥ 0.8"), (0.8, None))

    def test_both_bounds_from_words(self):
        self.assertEqual(parse_spec("0.5 이상 1.5 이하"), (0.5, 1.5))

    def test_float_noise_does_not_leak(self):
        """공차 연산의 부동소수점 오차가 한계값에 남으면 상·하한 측정값이 오판된다."""
        self.assertEqual(parse_spec("25.40±0.20"), (25.2, 25.6))
        self.assertEqual(parse_spec("12.70±0.15"), (12.55, 12.85))
        self.assertEqual(parse_spec("1.25+0.10/-0.05"), (1.2, 1.35))
        self.assertEqual(parse_spec("0.50+0.05/-0.03"), (0.47, 0.55))

    def test_non_numeric(self):
        self.assertEqual(parse_spec("스크래치·이물 없음"), (None, None))
        self.assertEqual(parse_spec(""), (None, None))
        self.assertEqual(parse_spec("1.25"), (None, None))

    def test_judge_type(self):
        self.assertEqual(guess_judge_type("1.25±0.1"), "numeric")
        self.assertEqual(guess_judge_type("MAX 2.0"), "numeric")
        self.assertEqual(guess_judge_type("외관 이상 없음"), "ox")


class TestJudgeValues(unittest.TestCase):
    def test_numeric_ok_ng(self):
        self.assertEqual(judge_values("numeric", 1.15, 1.35, ["1.2", "1.3"]), "OK")
        self.assertEqual(judge_values("numeric", 1.15, 1.35, ["1.2", "1.40"]), "NG")
        self.assertEqual(judge_values("numeric", None, 2.0, ["1.9"]), "OK")
        self.assertEqual(judge_values("numeric", 1.0, None, ["0.9"]), "NG")

    def test_numeric_edge_is_ok(self):
        self.assertEqual(judge_values("numeric", 1.15, 1.35, ["1.15", "1.35"]), "OK")

    def test_tolerance_edge_from_spec_is_ok(self):
        """규격 상·하한에 정확히 걸린 측정값은 합격이어야 한다."""
        lsl, usl = parse_spec("25.40±0.20")
        self.assertEqual(judge_values("numeric", lsl, usl, ["25.60"]), "OK")
        self.assertEqual(judge_values("numeric", lsl, usl, ["25.20"]), "OK")
        self.assertEqual(judge_values("numeric", lsl, usl, ["25.61"]), "NG")

    def test_blank_is_unjudged(self):
        self.assertEqual(judge_values("numeric", 1.0, 2.0, ["", "  "]), "")
        self.assertEqual(judge_values("ox", None, None, []), "")

    def test_partial_entry_judges_entered_only(self):
        self.assertEqual(judge_values("numeric", 1.0, 2.0, ["1.5", ""]), "OK")

    def test_unparseable_numeric_is_ng(self):
        self.assertEqual(judge_values("numeric", 1.0, 2.0, ["측정불가"]), "NG")

    def test_no_limits_cannot_autojudge(self):
        self.assertEqual(judge_values("numeric", None, None, ["1.5"]), "")

    def test_ox(self):
        self.assertEqual(judge_values("ox", None, None, ["OK", "양호"]), "OK")
        self.assertEqual(judge_values("ox", None, None, ["OK", "NG"]), "NG")
        self.assertEqual(judge_values("ox", None, None, ["OK", "애매"]), "")


if __name__ == "__main__":
    unittest.main()
