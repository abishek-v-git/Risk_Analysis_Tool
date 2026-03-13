from django.contrib import admin
from .models import ReportAnalysis

@admin.register(ReportAnalysis)
class ReportAnalysisAdmin(admin.ModelAdmin):
    list_display = ('filename', 'user', 'upload_date', 'high_risk_count', 'low_risk_count', 'no_risk_count')
    list_filter = ('user', 'upload_date')
    search_fields = ('filename', 'user__username')
