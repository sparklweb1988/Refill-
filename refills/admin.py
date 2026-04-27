from django.contrib import admin
from .models import Refill, Facility

# Unregister if already registered
if Facility in admin.site._registry:
    admin.site.unregister(Facility)
    
@admin.register(Facility)
class FacilityAdmin(admin.ModelAdmin):
    list_display = ("name",)
    search_fields = ("name",)


@admin.register(Refill)
class RefillAdmin(admin.ModelAdmin):
    list_display = (
        "unique_id",
        "facility",
        "sex",
        "last_pickup_date",
        "months_of_refill_days",
        "next_appointment",
        "case_manager",
    )
    list_filter = ("facility", "sex", "months_of_refill_days")
    search_fields = ("unique_id",)