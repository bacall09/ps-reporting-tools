# Compatibility shim — all content has moved to config.py at the root
# This file exists so shared/loaders.py, whs.py etc. can keep their imports unchanged
from config import *  # noqa: F401,F403
