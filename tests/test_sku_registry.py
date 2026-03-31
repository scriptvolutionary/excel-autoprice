from __future__ import annotations

import unittest

from seasonal_price.domain.services.sku_registry import (
    SkuRegistryService,
    normalize_category_code,
)


class SkuRegistryServiceTests(unittest.TestCase):
    def test_next_sku_is_unique_and_sequential(self) -> None:
        service = SkuRegistryService(existing_skus={"ROS-000001"})
        self.assertEqual(service.next_sku("ROS"), "ROS-000002")
        self.assertEqual(service.next_sku("ROS"), "ROS-000003")

    def test_validate_pattern(self) -> None:
        self.assertTrue(SkuRegistryService.validate("ABC-000123"))
        self.assertFalse(SkuRegistryService.validate("abc-123"))

    def test_normalize_category_code(self) -> None:
        self.assertEqual(normalize_category_code("Роза"), "РОЗ")
        self.assertEqual(normalize_category_code(""), "CAT")


if __name__ == "__main__":
    unittest.main()
