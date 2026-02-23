from django import forms
from .models import Refill, Facility
from django.utils import timezone
from datetime import timedelta
from django.core.exceptions import ValidationError

# -----------------------
# Refill Form (existing)
# -----------------------







# =====================
# Refill Form
# =====================
class RefillForm(forms.ModelForm):
    class Meta:
        model = Refill
        fields = [

            'facility',
            'unique_id',
            'art_start_date',                  # ART Start Date
            'vl_sample_collection_date',       # Viral Load Sample Collection
            'vl_result',                       # Viral Load Result
            'last_pickup_date',
            'sex',
            'months_of_refill_days',           # now decimal
            'current_regimen',
            'case_manager',
            'remark',

            'facility', 'unique_id', 'art_start_date', 'vl_sample_collection_date',
            'vl_result', 'last_pickup_date', 'sex', 'months_of_refill_days',
            'current_regimen', 'case_manager', 'remark',
            'tpt_start_date', 'tpt_completion_date'

        ]
        widgets = {
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'unique_id': forms.TextInput(attrs={'class': 'form-control'}),
            'art_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_result': forms.NumberInput(attrs={'class': 'form-control', 'placeholder': 'copies/ml'}),
            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),
            'months_of_refill_days': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.1'}),
            'current_regimen': forms.TextInput(attrs={'class': 'form-control'}),
            'case_manager': forms.TextInput(attrs={'class': 'form-control'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'tpt_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'tpt_completion_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        }

    def clean_vl_result(self):
        vl = self.cleaned_data.get('vl_result')
        if vl is not None and vl < 0:
            raise ValidationError("Viral Load cannot be negative.")
        return vl


    def save(self, commit=True):
        """
        Override save to automatically calculate:
        - VL eligibility (≥180 days on ART)
        - VL status is now read-only in model, so we DO NOT assign it here
        """
        instance = super().save(commit=False)

        today = timezone.now().date()

        # Calculate VL eligibility
        instance.vl_eligible = False
        if instance.art_start_date and (today - instance.art_start_date).days >= 180:
            instance.vl_eligible = True

        # VL status is read-only property, so remove any assignments
        # Determine current quarter if needed elsewhere
        def get_quarter(date):
            if not date:
                return None
            month = date.month
            if month in [1, 2, 3]:
                return "Q1"
            elif month in [4, 5, 6]:
                return "Q2"
            elif month in [7, 8, 9]:
                return "Q3"
            else:
                return "Q4"

        current_quarter = get_quarter(today)

        # Save instance
        if commit:
            instance.save()
        return instance


    def clean(self):
        cleaned_data = super().clean()
        art_date = cleaned_data.get("art_start_date")
        vl_date = cleaned_data.get("vl_sample_collection_date")
        if art_date and vl_date and vl_date < art_date:
            raise ValidationError("VL sample date cannot be before ART start date.")
        return cleaned_data
    
    
    

class UploadExcelForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )

