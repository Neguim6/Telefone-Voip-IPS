#!/usr/bin/env python3
"""Entry point para executar via `python run.py`"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from voip_cli.app import main

if __name__ == "__main__":
    main()
