"""
Django models mirroring the Phase 1 PostgreSQL schema (see
research/phase1/db.py). Field names, nullability and indexes match the raw
SQL so the web layer reads/writes the same shape the researcher pipeline does.

Numeric mapping: money-like columns -> DecimalField; scores / ratios / trends
(unconstrained NUMERIC in the raw schema) -> FloatField. The denormalised
`asin` links on history/insights are kept as plain text (not ForeignKeys) to
match the original design.
"""

from django.db import models


class Product(models.Model):
    asin = models.CharField(max_length=20, primary_key=True)
    title = models.TextField(null=True, blank=True)
    brand = models.CharField(max_length=255, null=True, blank=True)
    category = models.CharField(max_length=255, null=True, blank=True)
    subcategory = models.CharField(max_length=255, null=True, blank=True)

    price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    bsr = models.IntegerField(null=True, blank=True)
    estimated_sales = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)   # ASIN-level monthly sales
    estimated_revenue = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)  # ASIN-level monthly revenue
    review_count = models.IntegerField(null=True, blank=True)
    rating = models.FloatField(null=True, blank=True)
    weight_kg = models.FloatField(null=True, blank=True)
    yoy_growth = models.FloatField(null=True, blank=True)

    is_seasonal = models.BooleanField(null=True, blank=True)
    variation_count = models.IntegerField(null=True, blank=True)
    has_variants = models.BooleanField(null=True, blank=True)
    is_amazon_seller = models.BooleanField(null=True, blank=True)
    is_compatibility = models.BooleanField(null=True, blank=True)
    is_global_brand = models.BooleanField(null=True, blank=True)

    fulfillment = models.CharField(max_length=50, null=True, blank=True)
    image_url = models.TextField(null=True, blank=True)
    url = models.TextField(null=True, blank=True)
    listing_age_months = models.FloatField(null=True, blank=True)
    sales_trend_90d = models.FloatField(null=True, blank=True)
    price_trend_90d = models.FloatField(null=True, blank=True)

    price_range = models.CharField(max_length=10, null=True, blank=True)     # low / mid / high
    review_bucket = models.CharField(max_length=10, null=True, blank=True)   # low / medium / high
    sales_bucket = models.CharField(max_length=10, null=True, blank=True)    # low / medium / high

    competition_score = models.FloatField(null=True, blank=True)
    demand_score = models.FloatField(null=True, blank=True)
    opportunity_score = models.FloatField(null=True, blank=True)

    llm_adj = models.FloatField(null=True, blank=True)                       # Claude judgement adjustment (-10 to +10)
    llm_reason = models.TextField(null=True, blank=True)                     # Claude's one-line reason
    llm_at = models.DateTimeField(null=True, blank=True)                     # when Claude scoring was applied

    created_at = models.DateTimeField(auto_now_add=True)
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = "products"
        managed = False  # table is created/owned by research/phase1/db.py
        indexes = [
            models.Index(fields=["category"], name="idx_products_category"),
            models.Index(fields=["price_range"], name="idx_products_price_range"),
            models.Index(fields=["-opportunity_score"], name="idx_products_opportunity"),
            models.Index(fields=["has_variants"], name="idx_products_has_variants"),
            models.Index(fields=["is_seasonal"], name="idx_products_is_seasonal"),
        ]

    def __str__(self):
        return f"{self.asin} — {self.title or ''}".strip(" —")


class ProductHistory(models.Model):
    asin = models.CharField(max_length=20, db_index=True)
    price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    bsr = models.IntegerField(null=True, blank=True)
    estimated_sales = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    review_count = models.IntegerField(null=True, blank=True)
    captured_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "product_history"
        managed = False  # table is created/owned by research/phase1/db.py
        indexes = [
            models.Index(fields=["asin"], name="idx_history_asin"),
        ]

    def __str__(self):
        return f"{self.asin} @ {self.captured_at:%Y-%m-%d}"


class ProductReviewsInsights(models.Model):
    """Claude's distilled review analysis for one product (one row per ASIN).

    Mirrors the `ReviewAnalysis` Claude returns in
    research/phase3/claude_client.py: top-3 weaknesses, an index-aligned
    solution per weakness, and top-3 strengths. This is the cached *verdict*
    over the raw rows in `product_reviews`. Django-managed (the table is part
    of the new web layer, not the legacy ingest schema in phase1/db.py).

    `updated_at` is the freshness marker for re-scrape decisions (e.g. reuse
    if younger than N days, else re-analyse).
    """

    product = models.OneToOneField(
        Product,
        to_field="asin",
        db_column="asin",
        on_delete=models.CASCADE,
        related_name="insight",
        primary_key=True,
    )
    weaknesses = models.JSONField(default=list, blank=True)   # top 3 recurring complaints (1-3★)
    solutions = models.JSONField(default=list, blank=True)    # concrete fix per weakness, index-aligned
    strengths = models.JSONField(default=list, blank=True)    # top 3 praised features (4-5★)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = "product_reviews_insights"
        verbose_name_plural = "Product reviews insights"

    def __str__(self):
        return f"insight:{self.product_id}"


class Review(models.Model):
    """One individual scraped Amazon review, linked to its product by ASIN.

    Unlike `products` (owned by research/phase1/db.py), this table is created
    and owned by Django (managed=True), so `migrate` builds it. The FK column
    is literally `asin` and references products.asin (the Product PK). The
    fields mirror the normalised review shape from
    research/phase3/reviews_flyby.py (_norm: stars/title/body/verified/date).
    """

    product = models.ForeignKey(
        Product,
        to_field="asin",
        db_column="asin",
        on_delete=models.CASCADE,
        related_name="reviews",
    )
    review_id = models.CharField(max_length=255)            # Amazon review_id, or synthetic dedup key
    stars = models.IntegerField(null=True, blank=True)
    title = models.TextField(null=True, blank=True)
    body = models.TextField(null=True, blank=True)
    verified = models.BooleanField(null=True, blank=True)   # is_verified_purchase
    review_date = models.CharField(max_length=255, null=True, blank=True)  # free-form API date string
    marketplace = models.CharField(max_length=10, null=True, blank=True)
    scraped_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "product_reviews"
        constraints = [
            # One row per (product, review) — lets re-scrapes upsert/skip dupes.
            models.UniqueConstraint(fields=["product", "review_id"], name="uq_review_per_product"),
        ]
        indexes = [
            models.Index(fields=["product"], name="idx_reviews_asin"),
            models.Index(fields=["stars"], name="idx_reviews_stars"),
        ]

    def __str__(self):
        return f"{self.product_id} [{self.stars}*] {self.title or ''}".strip()
