"""Shared helpers for generating validation-safe list values in examples."""

from __future__ import annotations

import random
from math import ceil, floor, isfinite
from numbers import Real
from typing import Any


_MAX_REASONABLE_SHAREPOINT_BOUND = 1e50


def coerce_numeric_bound(value: Any) -> float | None:
    """Normalize numeric bounds and ignore unusable SharePoint sentinel values."""
    if value is None:
        return None

    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return None

    if not isfinite(numeric_value):
        return None

    if abs(numeric_value) > _MAX_REASONABLE_SHAREPOINT_BOUND:
        return None

    return numeric_value


def random_number_from_validation(validation: dict[str, Any]) -> int | float:
    """Generate a random number that honors min/max and decimal precision rules."""
    minimum = coerce_numeric_bound(validation.get("minimum"))
    maximum = coerce_numeric_bound(validation.get("maximum"))
    decimal_places = _resolve_decimal_places(validation, minimum, maximum)

    if minimum is None and maximum is None:
        return random.randint(0, 10000)

    if minimum is None:
        assert maximum is not None
        minimum = maximum - 1000
    if maximum is None:
        assert minimum is not None
        maximum = minimum + 1000

    if minimum > maximum:
        minimum, maximum = maximum, minimum

    if decimal_places is not None and decimal_places > 0:
        value = round(random.uniform(minimum, maximum), decimal_places)
        return _clamp(value, minimum, maximum)

    int_min = ceil(minimum)
    int_max = floor(maximum)
    if int_min <= int_max:
        return random.randint(int_min, int_max)

    # If the range contains no integer values (for example 0.1..0.9),
    # return a bounded float that still satisfies min/max.
    value = round(random.uniform(minimum, maximum), 2)
    return _clamp(value, minimum, maximum)


def bounded_number_update(
    current_value: Any,
    validation: dict[str, Any],
    increment: float,
) -> int | float:
    """Return an updated numeric value that always respects validation limits."""
    minimum = coerce_numeric_bound(validation.get("minimum"))
    maximum = coerce_numeric_bound(validation.get("maximum"))
    decimal_places = _resolve_decimal_places(validation, minimum, maximum)

    if isinstance(current_value, Real) and not isinstance(current_value, bool):
        candidate = float(current_value) + float(increment)
    else:
        return random_number_from_validation(validation)

    if minimum is not None or maximum is not None:
        candidate = _clamp(candidate, minimum, maximum)

    if decimal_places is not None and decimal_places > 0:
        candidate = round(candidate, decimal_places)
        if minimum is not None or maximum is not None:
            candidate = _clamp(candidate, minimum, maximum)
        return candidate

    return candidate


def _resolve_decimal_places(
    validation: dict[str, Any],
    minimum: float | None,
    maximum: float | None,
) -> int | None:
    """Map SharePoint-style decimal metadata to an integer decimal precision."""
    decimal_places_raw = validation.get("decimal_places")

    if isinstance(decimal_places_raw, int):
        return min(2, max(0, decimal_places_raw))

    if isinstance(decimal_places_raw, str):
        normalized = decimal_places_raw.strip().casefold()
        if normalized in {"none", "0", "zero"}:
            return 0
        if normalized in {"one", "1"}:
            return 1
        if normalized in {"two", "2"}:
            return 2
        if normalized in {"automatic", "auto"}:
            return (
                2
                if _has_fractional_part(minimum) or _has_fractional_part(maximum)
                else 0
            )

    if _has_fractional_part(minimum) or _has_fractional_part(maximum):
        return 2

    return None


def _has_fractional_part(value: float | None) -> bool:
    return value is not None and not value.is_integer()


def _clamp(value: float, minimum: float | None, maximum: float | None) -> float:
    if minimum is not None and value < minimum:
        value = minimum
    if maximum is not None and value > maximum:
        value = maximum
    return value
