from django.db import models
from datetime import timedelta
from django.utils import timezone
from decimal import Decimal
from django.db.models import F, Q


class Facility(models.Model):
    name = models.CharField(max_length=255, unique=True)
    code = models.CharField(max_length=50, unique=True)
    location = models.CharField(max_length=255, blank=True, null=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name








class Refill(models.Model):
    SEX_CHOICES = (('Male', 'Male'), ('Female', 'Female'))
    STATUS_CHOICES = (('Active', 'Active'), ('Active Restart', 'Active Restart'), ('Inactive', 'Inactive'))

    facility = models.ForeignKey(Facility, on_delete=models.CASCADE, related_name="refills")
    unique_id = models.CharField(max_length=100)
    last_pickup_date = models.DateField(null=True, blank=True)
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)
    months_of_refill_days = models.DecimalField(max_digits=4, decimal_places=2)
    current_regimen = models.CharField(max_length=255)
    case_manager = models.CharField(max_length=255)
    remark = models.TextField(blank=True, null=True)
    current_art_status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='Active')
    next_appointment = models.DateField(blank=True, null=True)
    expected_iit_date = models.DateField(blank=True, null=True)
    missed_appointment = models.BooleanField(default=False)

    # Viral Load
    art_start_date = models.DateField(blank=True, null=True)
    vl_sample_collection_date = models.DateField(blank=True, null=True)
    vl_result = models.IntegerField(blank=True, null=True)  # copies/ml

    # TPT
    tpt_start_date = models.DateField(blank=True, null=True)
    tpt_completion_date = models.DateField(blank=True, null=True)
    tpt_expected_completion = models.DateField(blank=True, null=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('facility', 'unique_id')
        ordering = ['-last_pickup_date']

    def calculate_dates(self):
        if self.last_pickup_date and self.months_of_refill_days:
            days = float(self.months_of_refill_days) * 30
            self.next_appointment = self.last_pickup_date + timedelta(days=days)
            self.expected_iit_date = self.next_appointment + timedelta(days=28)

    def save(self, *args, **kwargs):
        today = timezone.now().date()
        self.calculate_dates()
        if self.next_appointment and self.next_appointment < today:
            self.missed_appointment = True

        if self.tpt_start_date:
            self.tpt_expected_completion = self.tpt_start_date + timedelta(days=180)
        else:
            self.tpt_expected_completion = None

        super().save(*args, **kwargs)

    @property
    def is_vl_eligible(self):
        if not self.art_start_date:
            return False
        days_on_art = (timezone.now().date() - self.art_start_date).days
        if days_on_art < 180:
            return False
        age_years = days_on_art // 365
        vl_date = self.vl_sample_collection_date
        if age_years >= 15:
            return not (vl_date and vl_date.year == timezone.now().year)
        else:
            return not (vl_date and (timezone.now().date() - vl_date).days < 180)

    @property
    def is_suppressed(self):
        if self.vl_result is None:
            return None
        return self.vl_result < 1000

    @property
    def vl_status(self):
        if self.is_vl_eligible:
            return "Eligible"
        return "Not Eligible"

    @property
    def tpt_status(self):
        if not self.tpt_start_date:
            return "Not Started"
        if self.tpt_completion_date:
            return "Completed"
        if self.tpt_expected_completion and timezone.now().date() > self.tpt_expected_completion:
            return "Overdue"
        return "Ongoing"

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"