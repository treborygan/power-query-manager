from .models import ExistingQueryMode


def parse_existing_query_mode(value) -> ExistingQueryMode:
    """Normalize Qt combo-box user data to ExistingQueryMode.

    PySide may return a str for str-based Enum values after packaging,
    so never assume the returned object still has a .value attribute.
    """
    if isinstance(value, ExistingQueryMode):
        return value
    return ExistingQueryMode(str(value))
