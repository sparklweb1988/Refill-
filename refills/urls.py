from django.urls import path, re_path  # Import re_path for regex matching
from . import views
urlpatterns = [
    path('', views.signin_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('refills/', views.refill_list, name='refill_list'),
    path('refills/add/', views.refill_create, name='refill_add'),
    path('refills/edit/<int:pk>/', views.refill_update, name='refill_edit'),
    path('upload/', views.upload_excel, name='upload_excel'),
    path('refills/daily/', views.daily_refill_list, name='daily_refill_list'),
    path('refills/track/', views.track_refills, name='track_refills'),
    
    # Update URL pattern for 'refill_add_with_id' to handle unique_id with slashes
    re_path(r'^refills/add/(?P<unique_id>.+)/$', views.refill_create, name='refill_add_with_id'),  
    path('missed-refills/', views.missed_refills, name='missed_refills'),
    path('track-vl/', views.track_vl, name='track_vl'),
]

