"""
Phase 1 is now a Django ORM library, not a CLI. Querying the stored dataset
happens through the web layer:

    from research.phase1 import run_query
    result = run_query(budget=2000, category=None, include_seasonal=True)

Phases 2 and 3 are likewise libraries now (research.services.phase2.run_phase2,
research.services.phase3.run_phase3). Run the site with
`python manage.py runserver`.
"""

if __name__ == "__main__":
    print(__doc__)
