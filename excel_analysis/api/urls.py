from django.urls import path
from .views import ExcelAnalysisView

urlpatterns = [
    path("analyze/", ExcelAnalysisView.as_view(), name="analyze_excel"),
]
