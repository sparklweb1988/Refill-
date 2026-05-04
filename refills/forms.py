from django import forms
from .models import Refill
from datetime import timedelta


# ================= MISSED REASON DROPDOWN =================
MISSED_REASON_CHOICES = [
    ("", "Select reason"),
    ("TRAVEL", "Travel"),
    ("TRANSFER", "Transferred Out"),
    ("LOSS_TO_FOLLOW_UP", "Lost to Follow-Up"),
    ("SIDE_EFFECTS", "Side Effects"),
    ("STOCK_OUT", "Drug Stock Out"),
    ("FINANCIAL", "Financial Constraints"),
    ("FEELING_WELL", "Feeling Well / Stopped Care"),
    ("DEATH", "Death"),
    ("OTHER", "Other"),
]


class RefillForm(forms.ModelForm):

    # override field explicitly for dropdown
    missed_reason = forms.ChoiceField(
        choices=MISSED_REASON_CHOICES,
        required=False,
        widget=forms.Select(attrs={'class': 'form-select'})
    )

    class Meta:
        model = Refill
        fields = '__all__'

        widgets = {

            # ================= PATIENT INFO =================
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'unique_id': forms.TextInput(attrs={'class': 'form-control'}),
            'age': forms.NumberInput(attrs={'class': 'form-control'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),

            # ================= ART & VL =================
            'art_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_result': forms.NumberInput(attrs={'class': 'form-control'}),

            # ================= REFILL =================
            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'months_of_refill_days': forms.NumberInput(attrs={'class': 'form-control'}),
            'current_regimen': forms.TextInput(attrs={'class': 'form-control'}),
            'case_manager': forms.TextInput(attrs={'class': 'form-control'}),

            # ================= TB =================
            'tb_screening_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_screening_type': forms.Select(attrs={'class': 'form-select'}),
            'tb_status': forms.Select(attrs={'class': 'form-select'}),
            'tb_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_result_received_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_diagnostic_result': forms.Select(attrs={'class': 'form-select'}),

            # ================= TRACKING =================
            'tracking_date_1': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracking_date_2': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracking_date_3': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracked_by': forms.TextInput(attrs={'class': 'form-control'}),

            # ================= DISCONTINUATION =================
            'patient_discontinued': forms.Select(attrs={'class': 'form-select'}),
            'discontinued_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'returned_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),

            # IMPORTANT: we removed discontinued_reason widget override because dropdown is handled above

            # ================= NOTES =================
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),

            # ================= TPT =================
            'tpt_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tpt_completion_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),

            # ================= EAC =================
            'eac_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        }

    # ================= VALIDATION =================
    def clean(self):
        cleaned_data = super().clean()

        age = cleaned_data.get("age")
        art = cleaned_data.get("art_start_date")
        vl_date = cleaned_data.get("vl_sample_collection_date")
        vl_result = cleaned_data.get("vl_result")

        discontinued = cleaned_data.get("patient_discontinued")
        discontinued_date = cleaned_data.get("discontinued_date")
        missed_reason = cleaned_data.get("missed_reason")
        remark = cleaned_data.get("remark")

        # -------- AGE CHECK --------
        if age is not None and age < 0:
            self.add_error("age", "Age cannot be negative.")

        # -------- VL LOGIC --------
        if art and vl_date and vl_date < art:
            self.add_error("vl_sample_collection_date", "VL date cannot be before ART start date.")

        if art and vl_date:
            six_months = art + timedelta(days=180)

            is_first_vl = (
                not self.instance.pk
                or not self.instance.vl_sample_collection_date
            )

            if is_first_vl and vl_date < six_months:
                self.add_error(
                    "vl_sample_collection_date",
                    "First VL should not be before 6 months on ART."
                )

        if vl_result is not None and not vl_date:
            self.add_error("vl_sample_collection_date", "Enter VL sample collection date first.")

        if vl_date and not art:
            self.add_error("art_start_date", "ART start date is required before VL entry.")

        # -------- DISCONTINUATION LOGIC --------
        if discontinued == "Y":

            # must have date
            if not discontinued_date:
                self.add_error("discontinued_date", "Discontinued date is required when patient is marked as discontinued.")

            # missed reason required
            if not missed_reason:
                self.add_error("missed_reason", "Please select a reason for discontinuation.")

            # remark should support reasoning
            if not remark:
                self.add_error("remark", "Please add a note/remark for discontinuation context.")

        else:
            # if NOT discontinued → clear irrelevant fields
            cleaned_data["discontinued_date"] = None
            cleaned_data["missed_reason"] = None

        return cleaned_data


# ================= EXCEL UPLOAD =================
class UploadExcelForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )