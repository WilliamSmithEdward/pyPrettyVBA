"""Test configuration: shared imports and Hypothesis profiles.

tests/ is put on the import path for the shared fixture support and
strategies. HYPOTHESIS_PROFILE picks a Hypothesis profile: `default` for a
quick run, `thorough` for a long one.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from hypothesis import HealthCheck, settings

sys.path.insert(0, str(Path(__file__).resolve().parent))

settings.register_profile(
    "default", max_examples=100, deadline=None, suppress_health_check=[HealthCheck.too_slow]
)
settings.register_profile(
    "thorough", max_examples=3000, deadline=None, suppress_health_check=[HealthCheck.too_slow]
)
settings.load_profile(os.environ.get("HYPOTHESIS_PROFILE", "default"))
