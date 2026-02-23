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

