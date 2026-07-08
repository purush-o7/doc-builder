import sys
from . import build

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m doc_builder <spec.json>")
        sys.exit(1)
    build(sys.argv[1])
