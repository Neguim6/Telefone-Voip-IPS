#!/usr/bin/env python3
"""Entry point for production (gunicorn)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from voip_cli.web import app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
