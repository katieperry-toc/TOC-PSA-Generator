import os
import sys
_APP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Champion TOC PSA Generator")
sys.path.insert(0, _APP_DIR)
from app import main
main()
