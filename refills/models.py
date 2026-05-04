from django.db import models
from datetime import timedelta
from django.utils import timezone
from decimal import Decimal
from django.db.models import F, Q
from django.utils import timezone
from dateutil.relativedelta import relativedelta
from datetime import date

class Facility(models.Model):
    name = models.CharField(max_length=255, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name









class Refill(models.Model):

    SEX_CHOICES = (
        ('Male', 'Male'),
        ('Female', 'Female')
    )

    STATUS_CHOICES = (
        ('Active', 'Active'),
        ('Active Restart', 'Active Restart'),
        ('Restart', 'Restart'),
        ('Inactive', 'Inactive')
    )

    TB_SCREENING_TYPE_CHOICES = (
        ('Symptom Screening', 'Symptom Screening'),
        ('Chest X-ray', 'Chest X-ray'),
        ('GeneXpert', 'GeneXpert'),
        ('LAM', 'LAM')
    )

    TB_STATUS_CHOICES = (
        ('No TB Symptoms', 'No TB Symptoms'),
        ('Presumptive TB', 'Presumptive TB'),
        ('TB Confirmed', 'TB Confirmed')
    )

    TB_RESULT_CHOICES = (
        ('Positive', 'Positive'),
        ('Negative', 'Negative'),
        ('Indeterminate', 'Indeterminate')
    )

    TB_CASCADE_CHOICES = (
        ('Presumptive', 'Presumptive'),
        ('Confirmed', 'Confirmed'),
        ('Negative', 'Negative'),
    )

    YES_NO_CHOICES = (
        ('Y', 'Yes'),
        ('N', 'No')
    )

    # ================= FACILITY =================
    facility = models.ForeignKey(
        "Facility",
        on_delete=models.CASCADE,
        related_name="refills"
    )

    tb_cascade_status = models.CharField(
        max_length=50,
        choices=TB_CASCADE_CHOICES,
        blank=True,
        null=True
    )
    
    missed_reason = models.CharField(
    max_length=255,
    blank=True,
    null=True
    )
    
    
    


    # ================= BASIC INFO =================
    unique_id = models.CharField(max_length=100)
    
    def save(self, *args, **kwargs):
        if self.unique_id:
            self.unique_id = self.unique_id.strip()
        super().save(*args, **kwargs)
        
        
    age = models.IntegerField(null=True, blank=True)
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)

    last_pickup_date = models.DateField(null=True, blank=True)
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

    # ================= TB =================
    tb_screening_date = models.DateField(blank=True, null=True)
    tb_screening_type = models.CharField(max_length=50, choices=TB_SCREENING_TYPE_CHOICES, blank=True, null=True)
    tb_status = models.CharField(max_length=50, choices=TB_STATUS_CHOICES, blank=True, null=True)
    tb_sample_collection_date = models.DateField(blank=True, null=True)
    tb_result_received_date = models.DateField(blank=True, null=True)
    tb_diagnostic_result = models.CharField(max_length=50, choices=TB_RESULT_CHOICES, blank=True, null=True)

    # ================= TRACKING =================
    tracking_date_1 = models.DateField(null=True, blank=True)
    tracking_date_2 = models.DateField(null=True, blank=True)
    tracking_date_3 = models.DateField(null=True, blank=True)

    tracked_by = models.CharField(max_length=100, null=True, blank=True)

    # ================= DISCONTINUATION =================
    patient_discontinued = models.CharField(
        max_length=1,
        choices=YES_NO_CHOICES,
        null=True,
        blank=True
    )

    discontinued_reason = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    discontinued_date = models.DateField(null=True, blank=True)
    returned_date = models.DateField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    # ================= META =================
    class Meta:
        ordering = ['facility__name', '-last_pickup_date']
    # ================= AUTO CALC =================
    def calculate_dates(self):
        if self.last_pickup_date and self.months_of_refill_days:
            self.next_appointment = self.last_pickup_date + timedelta(
                days=float(self.months_of_refill_days) * 30
            )
            self.expected_iit_date = self.next_appointment + timedelta(days=28)

    def save(self, *args, **kwargs):
        self.calculate_dates()

        if self.next_appointment and self.next_appointment < timezone.now().date():
            self.missed_appointment = True

        if self.tpt_start_date:
            self.tpt_expected_completion = self.tpt_start_date + timedelta(days=180)

        super().save(*args, **kwargs)

    # ================= SAFE DAYS MISSED =================
    @property
    def days_missed(self):
        if not self.next_appointment:
            return 0
        return max((timezone.now().date() - self.next_appointment).days, 0)
        # ================= VL CORE ENGINE =================

    @property
    def is_active_client(self):
        return self.current_art_status in ["Active", "Active Restart", "Restart"]


    @property
    def months_on_art(self):
        if not self.art_start_date:
            return 0
        today = timezone.now().date()
        return (today.year - self.art_start_date.year) * 12 + (today.month - self.art_start_date.month)


    @property
    def is_child(self):
        return self.age is not None and self.age < 15


    @property
    def is_adult(self):
        return not self.is_child


    # ================= LAST VL DATE =================
    @property
    def last_vl_date(self):
        return self.vl_sample_collection_date


    # ================= DUE DATE ENGINE =================
    @property
    def vl_due_date(self):
        """
        Core VL due calculation (NO hardcoding)
        """

        if not self.art_start_date:
            return None

        # First VL (after ART start)
        if not self.last_vl_date:
            return self.art_start_date + relativedelta(months=6)

        # Subsequent VL
        if self.is_child:
            return self.last_vl_date + relativedelta(months=6)
        else:
            return self.last_vl_date + relativedelta(months=12)


    # ================= QUARTER ENGINE =================
    @property
    def current_quarter(self):
        today = timezone.now().date()

        if today.month <= 3:
            return (date(today.year, 1, 1), date(today.year, 3, 31))
        elif today.month <= 6:
            return (date(today.year, 4, 1), date(today.year, 6, 30))
        elif today.month <= 9:
            return (date(today.year, 7, 1), date(today.year, 9, 30))
        else:
            return (date(today.year, 10, 1), date(today.year, 12, 31))


    # ================= SAMPLE COUNT CONTROL =================
    @property
    def samples_this_year(self):
        """
        Since your model stores only ONE VL,
        this safely enforces your rule.
        """

        if not self.last_vl_date:
            return 0

        today = timezone.now().date()

        if self.last_vl_date.year == today.year:
            return 1

        return 0


    # ================= CLINICAL ELIGIBILITY =================
    @property
    def is_vl_clinically_eligible(self):
        """
        PURE CLINICAL RULES ONLY (no quarter here)
        """

        # Active
        if not self.is_active_client:
            return False

        # Not discontinued
        if self.patient_discontinued == "Y":
            return False

        # ≥6 months on ART
        if self.months_on_art < 6:
            return False

        return True


    # ================= QUARTER ELIGIBILITY =================
    @property
    def is_vl_due_this_quarter(self):
        """
        Quarter rule:
        - Due in current quarter OR
        - Overdue from past quarter
        """

        if not self.is_vl_clinically_eligible:
            return False

        due_date = self.vl_due_date
        if not due_date:
            return False

        q_start, q_end = self.current_quarter

        # ✅ CRITICAL FIX: include overdue patients
        return due_date <= q_end


    # ================= FINAL PROGRAM ELIGIBILITY =================
    @property
    def is_vl_eligible_program(self):
        """
        MASTER RULE ENGINE (FINAL)
        """

        # Must pass quarter rule
        if not self.is_vl_due_this_quarter:
            return False

        # ================= SAMPLE CONTROL =================
        samples = self.samples_this_year

        if self.is_child:
            if samples >= 2:
                return False
        else:
            if samples >= 1:
                return False

        # ================= DUPLICATE VL PREVENTION =================
        due_date = self.vl_due_date

        if self.last_vl_date and due_date and self.last_vl_date >= due_date:
            return False

        return True


    # ================= HUMAN READABLE STATUS =================
    @property
    def vl_status(self):

        if not self.is_vl_eligible_program:
            return "Not Eligible"

        due = self.vl_due_date
        today = timezone.now().date()

        if due and due < today:
            return "Overdue VL"

        return "Due This Quarter"


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

    # ================= EAC =================
    @property
    def eac(self):
        return (
            self.vl_result is not None and
            self.vl_result >= 1000 and
            self.eac_start_date is None
        )

    @property
    def eac_status(self):
        if not self.eac:
            return "Not Eligible"
        if self.eac_sessions_completed == 0:
            return "Eligible for 1st EAC"
        if self.eac_sessions_completed == 1:
            return "Eligible for 2nd EAC"
        if self.eac_sessions_completed == 2:
            return "Eligible for 3rd EAC"
        return "Post-EAC VL Due"

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

    # ================= FIX TRIM =================
    
   
  
        
        
        
        
    @property
    def clean_unique_id(self):
        return (self.unique_id or "").strip()
    
    
        
    @property
    def safe_unique_id(self):
        if not self.unique_id:
            return None
        value = self.unique_id.strip()
        return value if value else None  
            
    # ================= IIT =================
    @property
    def iit_status(self):
        if self.days_missed >= 28:
            return "IIT"
        if self.days_missed > 0:
            return f"{28 - self.days_missed} days to IIT"
        return "On Track"

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"