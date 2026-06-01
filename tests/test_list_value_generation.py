"""Tests for shared list value generation helpers used by examples."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "examples"))

from list_value_generation import (
    bounded_number_update,
    random_number_from_validation,
)


def test_random_number_from_validation_within_bounds() -> None:
    """Random numbers should always respect configured numeric limits."""
    validation = {"minimum": -75.0, "maximum": -30.0, "decimal_places": "automatic"}

    for _ in range(100):
        value = float(random_number_from_validation(validation))
        assert -75.0 <= value <= -30.0


def test_bounded_number_update_clamps_to_maximum() -> None:
    """Incremented value should be clamped when it exceeds maximum."""
    validation = {"minimum": -75.0, "maximum": -30.0, "decimal_places": "automatic"}

    updated = float(
        bounded_number_update(
            current_value=-31.0,
            validation=validation,
            increment=10.0,
        )
    )

    assert updated == -30.0


def test_bounded_number_update_clamps_to_minimum() -> None:
    """Incremented value should be clamped when it goes below minimum."""
    validation = {"minimum": 0.0, "maximum": 1000.0, "decimal_places": "automatic"}

    updated = float(
        bounded_number_update(
            current_value=5.0,
            validation=validation,
            increment=-10.0,
        )
    )

    assert updated == 0.0


def test_bounded_number_update_fallback_respects_bounds() -> None:
    """Missing current values should use a bounded random fallback."""
    validation = {"minimum": 0.0, "maximum": 1000.0, "decimal_places": "none"}

    for _ in range(100):
        updated = float(
            bounded_number_update(
                current_value=None,
                validation=validation,
                increment=10.0,
            )
        )
        assert 0.0 <= updated <= 1000.0
