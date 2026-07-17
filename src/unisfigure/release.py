"""Central release-info module — single source of truth for install/release date."""
from . import RELEASE_DATE, __version__


def release_info() -> dict:
    return {
        "version": __version__,
        "release_date": RELEASE_DATE,
    }
