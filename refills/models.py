from django.db import models
from datetime import timedelta
from django.utils import timezone
from decimal import Decimal
from django.db.models import F, Q
from django.utils import timezone
from dateutil.relativedelta import relativedelta



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
    STATUS_CHOICES = (
        ('Active', 'Active'),
        ('Active Restart', 'Active Restart'),
        ('Restart', 'Restart'),
        ('Inactive', 'Inactive')
    )

    facility = models.ForeignKey("Facility", on_delete=models.CASCADE, related_name="refills")
    unique_id = models.CharField(max_length=100)
    last_pickup_date = models.DateField(null=True, blank=True)
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)
    months_of_refill_days = models.DecimalField(max_digits=4, decimal_places=2)
    current_regimen = models.CharField(max_length=255)
    case_manager = models.CharField(max_length=255)
    remark = models.TextField(blank=True, null=True)

    current_art_status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default='Active'
    )

    next_appointment = models.DateField(blank=True, null=True)
    expected_iit_date = models.DateField(blank=True, null=True)
    missed_appointment = models.BooleanField(default=False)

    # ================= VL =================
    art_start_date = models.DateField(blank=True, null=True)
    vl_sample_collection_date = models.DateField(blank=True, null=True)
    vl_result = models.IntegerField(blank=True, null=True)

    # ================= TPT =================
    tpt_start_date = models.DateField(blank=True, null=True)
    tpt_completion_date = models.DateField(blank=True, null=True)
    tpt_expected_completion = models.DateField(blank=True, null=True)

    # ================= EAC =================
    eac_start_date = models.DateField(blank=True, null=True)
    eac_sessions_completed = models.IntegerField(default=0)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('facility', 'unique_id')
        ordering = ['-last_pickup_date']

    # ================= AUTO DATES =================
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

    # ================= VL ELIGIBILITY =================
    @property
    def is_vl_eligible(self):
        today = timezone.now().date()
        if not self.art_start_date:
            return False
        if self.art_start_date + relativedelta(months=6) > today:
            return False
        if not self.vl_sample_collection_date:
            return True
        return self.vl_sample_collection_date + relativedelta(months=12) <= today

    # ================= SUPPRESSION =================
    @property
    def is_suppressed(self):
        if self.vl_result is None:
            return None
        return self.vl_result < 1000

    # ================= AHD =================
    @property
    def ahd(self):
        return self.current_art_status in ["Restart", "Active Restart"]

    # ================= EAC (MODEL-DRIVEN) =================
    @property
    def eac(self):
        return (
            self.vl_result is not None
            and self.vl_result >= 1000
            and self.vl_sample_collection_date is not None
            and self.eac_start_date is None
        )

    @property
    def eac_status(self):
        if not self.eac:
            return "Not Eligible"
        if self.eac_sessions_completed == 0:
            return "Eligible for 1st EAC"
        elif self.eac_sessions_completed == 1:
            return "Eligible for 2nd EAC"
        elif self.eac_sessions_completed == 2:
            return "Eligible for 3rd EAC"
        return "Post-EAC VL Due"

    @property
    def post_eac_vl_due(self):
        return self.eac and self.eac_sessions_completed >= 3

    # ================= TPT =================
    @property
    def tpt_status(self):
        today = timezone.now().date()
        if not self.tpt_start_date:
            return "Not Started"
        if self.tpt_completion_date:
            return "Completed"
        if self.tpt_expected_completion and today > self.tpt_expected_completion:
            return "Overdue"
        return "Ongoing"

    # ================= VL STATUS =================
    @property
    def vl_status(self):
        if not self.art_start_date and self.vl_result is None:
            return "Not Eligible"          # changed from "N/A"
        elif self.is_vl_eligible and self.vl_result is None:
            return "Eligible"              # changed from "Due" to match your wording
        elif self.is_suppressed:
            return "Suppressed"
        elif self.vl_result is not None:
            return "Unsuppressed"
        return "Not Eligible"              # fallback instead of "N/A"

    # ================= IIT STATUS =================
    @property
    def days_missed(self):
        if not self.next_appointment:
            return 0
        delta = (timezone.now().date() - self.next_appointment).days
        return max(delta, 0)

    @property
    def iit_status(self):
        if self.days_missed >= 28:
            return "IIT"
        elif 0 < self.days_missed < 28:
            return f"{28 - self.days_missed} days to IIT"
        return "On Track"

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"