from django.db import models
from django.contrib.auth.models import User

class ReportAnalysis(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='analyses')
    filename = models.CharField(max_length=255)
    upload_date = models.DateTimeField(auto_now_add=True)
    report_file_path = models.CharField(max_length=500)  # Path to the generated Excel report
    
    # Store the processed results as JSON for quick dashboard retrieval
    analysis_results = models.JSONField(null=True, blank=True)
    
    # Summary stats for quick filtering
    high_risk_count = models.IntegerField(default=0)
    low_risk_count = models.IntegerField(default=0)
    no_risk_count = models.IntegerField(default=0)

    class Meta:
        ordering = ['-upload_date']

    def __str__(self):
        return f"{self.filename} - {self.upload_date.strftime('%Y-%m-%d %H:%M')}"
