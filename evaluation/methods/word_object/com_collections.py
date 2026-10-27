"""com_collections with explicit evidence dependencies."""

from typing import List, Optional, Any


def strict_com_sequence(value: Any) -> Optional[List[Any]]:
    """Normalize scalar/SafeArray values without hiding COM failures."""
    if value is None:
        return None
    if isinstance(value, (tuple, list)):
        values = value
    else:
        values = [value]
    return [item.item() if hasattr(item, "item") else item for item in values]
