from django.contrib import admin
from .models import Refill, Facility, FacilityUser


# ================= FACILITY =================
# safe unregister (prevents duplicate crash)
try:
    admin.site.unregister(Facility)
except admin.sites.NotRegistered:
    pass


@admin.register(Facility)
class FacilityAdmin(admin.ModelAdmin):
    list_display = ("name",)
    search_fields = ("name",)


# ================= FACILITY USER =================
@admin.register(FacilityUser)
class FacilityUserAdmin(admin.ModelAdmin):
    list_display = ("user", "facility", "role", "created_at")
    list_filter = ("facility", "role")
    search_fields = ("user__username", "facility__name")


# ================= REFILL =================
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