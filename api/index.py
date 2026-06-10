import sys
import os

# Add the parent directory (project root) to python path so app can be imported
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
