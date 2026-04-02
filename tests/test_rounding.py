from __future__ import annotations

import unittest

from seasonal_price.domain.services.rounding import round_to_multiplicity


class RoundingServiceTests(unittest.TestCase):
    def test_round_down_mode(self) -> None:
        self.assertEqual(round_to_multiplicity(30, 24, mode="down"), (24, True))

    def test_round_up_mode(self) -> None:
        self.assertEqual(round_to_multiplicity(30, 24, mode="up"), (48, True))

    def test_round_up_mode_keeps_zero(self) -> None:
        self.assertEqual(round_to_multiplicity(0, 24, mode="up"), (0, False))

    def test_round_unknown_mode_raises(self) -> None:
        with self.assertRaises(ValueError):
            round_to_multiplicity(10, 5, mode="sideways")


if __name__ == "__main__":
    unittest.main()
