from django.urls import path
from django.contrib.auth import views as auth_views
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('api/table-data/', views.table_data_api, name='table_data_api'),
    path('api/analyze/', views.analyze_api, name='analyze_api'),
    path('login/', auth_views.LoginView.as_view(template_name='flagrisk/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
]
