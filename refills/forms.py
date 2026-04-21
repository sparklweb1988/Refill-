from django import forms
from .models import Refill, Facility
from django.core.exceptions import ValidationError
from datetime import timedelta

MISSED_REASONS = [
    ('', 'Select Reason'),
    ('Was Sick', 'Was Sick'),
    ('No Transport Fare', 'No Transport Fare'),
    ('Forgot', 'Forgot'),
    ('Felt Better', 'Felt Better'),
    ('Not Permitted to leave work', 'Not Permitted to leave work'),
    ('Last appointment Cancelled', 'Last appointment Cancelled'),
    ('Still had Drugs', 'Still had Drugs'),
    ('Taking Herbal Treatment', 'Taking Herbal Treatment'),
    ('Intense Followup', 'Intense Followup'),
]


class RefillForm(forms.ModelForm):
    missed_reason = forms.ChoiceField(
        choices=MISSED_REASONS,
        required=False,
        widget=forms.Select(attrs={'class': 'form-select'})
    )

    class Meta:
        model = Refill
        fields = [
            # Basic Info
            'facility', 'unique_id', 'age', 'sex',

            # ART & VL
            'art_start_date', 'vl_sample_collection_date', 'vl_result',

            # Refill
            'last_pickup_date', 'months_of_refill_days',
            'current_regimen', 'case_manager',

            # TB
            'tb_screening_date', 'tb_screening_type', 'tb_status',
            'tb_sample_collection_date', 'tb_result_received_date', 'tb_diagnostic_result',

            # Tracking
            'tracking_date_1', 'tracking_date_2', 'tracking_date_3',
            'tracked_by',

            # Discontinuation
            'patient_discontinued', 'discontinued_reason', 'discontinued_date',

            # Returned
            'returned_date',

            # Others
            'remark', 'tpt_start_date', 'tpt_completion_date',
        ]

        widgets = {
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'unique_id': forms.TextInput(attrs={'class': 'form-control'}),
            'age': forms.NumberInput(attrs={'class': 'form-control'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),

            'art_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_result': forms.NumberInput(attrs={'class': 'form-control'}),

            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'months_of_refill_days': forms.NumberInput(attrs={'class': 'form-control'}),

            'current_regimen': forms.TextInput(attrs={'class': 'form-control'}),
            'case_manager': forms.TextInput(attrs={'class': 'form-control'}),

            # TB
            'tb_screening_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_screening_type': forms.Select(attrs={'class': 'form-select'}),
            'tb_status': forms.Select(attrs={'class': 'form-select'}),
            'tb_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_result_received_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tb_diagnostic_result': forms.Select(attrs={'class': 'form-select'}),

            # Tracking
            'tracking_date_1': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracking_date_2': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracking_date_3': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tracked_by': forms.TextInput(attrs={'class': 'form-control'}),

            # Discontinued
            'patient_discontinued': forms.Select(attrs={'class': 'form-select'}),
            'discontinued_reason': forms.Select(attrs={'class': 'form-select'}),
            'discontinued_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),

            # Returned
            'returned_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),

            # Others
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'tpt_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tpt_completion_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        }

    # ================= VL VALIDATION =================
    def clean_vl_result(self):
        vl = self.cleaned_data.get('vl_result')

        if vl is None:
            return vl

        try:
            vl = int(vl)
        except (ValueError, TypeError):
            raise ValidationError("Viral Load must be a valid number.")

        if vl < 0:
            raise ValidationError("Viral Load cannot be negative.")

        if vl > 10000000:
            raise ValidationError("Viral Load value is unrealistically high.")

        return vl

    # ================= CLEAN (VL RULES FIXED) =================
    def clean(self):
        cleaned_data = super().clean()

        art_date = cleaned_data.get("art_start_date")
        vl_date = cleaned_data.get("vl_sample_collection_date")

        # ART vs VL consistency
        if art_date and vl_date:
            if vl_date < art_date:
                raise ValidationError("VL sample date cannot be before ART start date.")

            # 6-month ART eligibility rule
            if vl_date < art_date + timedelta(days=180):
                self.add_error(
                    "vl_sample_collection_date",
                    "VL should not be collected before 6 months on ART."
                )

        return cleaned_data

    # ================= SAVE =================
    def save(self, commit=True):
        instance = super().save(commit=False)

        missed_reason = self.cleaned_data.get('missed_reason')
        if missed_reason:
            instance.remark = missed_reason

        if commit:
            instance.save()

        return instance


class UploadExcelForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )