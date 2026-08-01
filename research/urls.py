from django.urls import path
from . import views

app_name = "research"

urlpatterns = [
    path("", views.research_dashboard, name="research"),
    path("products/", views.shortlist, name="products"),
    path("select/keywords/<str:asin>/", views.select_asin, name="select"),
    path("phase3/<str:asin>/", views.phase3, name="phase3"),
    path("phase4/<str:asin>/", views.phase4, name="phase4"),
    path("phase3/download/<str:name>/", views.download_workbook, name="download_workbook"),
]