"""
Root-level launcher. Always starts at Phase 1; Phase 1 hands off into Phase 2
(keyword verification) at the end, reusing the loaded Black Box data.
Usage:
    python run.py

(To exercise Phase 2 in isolation during development, run:
    python -m phase2.main
 — it will prompt for the Black Box file since there is no Phase 1 data.)
"""

from phase1.main import run_phase1

if __name__ == "__main__":
    run_phase1()
