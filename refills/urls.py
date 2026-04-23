from django.urls import path, re_path
from . import views

urlpatterns = [
    path('', views.signin_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('dashboard/', views.dashboard, name='dashboard'),

    # ================= REFILLS =================
    path('refills/', views.refill_list, name='refill_list'),
    path('refills/export/', views.export_refills_view, name='export_refills'),  # ✅ FIXED

    path('refills/add/', views.refill_create, name='refill_add'),
    path('refills/edit/<int:pk>/', views.refill_update, name='refill_edit'),

    path('upload/', views.upload_excel, name='upload_excel'),

    path('refills/daily/', views.daily_refill_list, name='daily_refill_list'),

    # ================= TRACK REFILLS =================
    path('refills/track/', views.track_refills, name='track_refills'),
    path('track-refills/export/', views.export_track_refills_view, name='export_track_refills'),  # ✅ FIXED

    # ================= MISSED REFILLS =================
    path('missed-refills/', views.missed_refills, name='missed_refills'),
    path('missed-refills/export/', views.export_missed_refills_view, name='export_missed_refills'),  # ✅ ADDED

    # ================= TRACK VL =================
    path('track-vl/', views.track_vl, name='track_vl'),
    path('track-vl/export/', views.export_vl_view, name='export_vl'),  # ✅ ADDED

    # ================= SPECIAL =================
    re_path(r'^refills/add/(?P<unique_id>.+)/$', views.refill_create, name='refill_add_with_id'),
]