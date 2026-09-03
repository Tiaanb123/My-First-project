"""PyInstaller entry point — keeps package-relative imports working when frozen."""
from reportgen.demo import main

if __name__ == "__main__":
    raise SystemExit(main())
