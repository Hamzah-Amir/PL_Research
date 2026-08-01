from django.contrib import admin
from .models import Product, ProductHistory, ProductReviewsInsights, Review

# Register your models here.

class ProductAdmin(admin.ModelAdmin):
    list_display = ("asin", "title", "category", "price", "is_seasonal")
    search_fields = ("asin", "title")
    list_filter = ("category", "is_seasonal")
admin.site.register(Product, ProductAdmin)
admin.site.register(ProductHistory)
admin.site.register(ProductReviewsInsights)
admin.site.register(Review)
