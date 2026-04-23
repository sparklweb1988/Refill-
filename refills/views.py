from dateutil.relativedelta import relativedelta
from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone
from django.db import transaction
from datetime import date, timedelta

from django.core.exceptions import ValidationError
from .forms import RefillForm, UploadExcelForm
from .models import Refill, Facility
import pandas as pd
import openpyxl
from django.http import HttpResponse
from django.conf import settings
from django.db.models import F, Q
from django.core.paginator import Paginator
from openpyxl import Workbook
from datetime import datetime

from django.contrib import messages
from openpyxl.styles import Font
from io import BytesIO

from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout





# views.py








def signin_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('pw')

        user = authenticate(username=username, password=password)
        if user is not None:
            login(request, user)
            messages.success(request, 'Login successful!')
            return redirect('dashboard')
    return render(request, 'signin.html')



def logout_view(request):
    logout(request)
    messages.success(request, ' Logout successfully')
    return redirect('login')








def get_quarter_range(today):
    if today.month in [1, 2, 3]:
        return date(today.year, 1, 1), date(today.year, 3, 31)
    elif today.month in [4, 5, 6]:
        return date(today.year, 4, 1), date(today.year, 6, 30)
    elif today.month in [7, 8, 9]:
        return date(today.year, 7, 1), date(today.year, 9, 30)
    else:
        return date(today.year, 10, 1), date(today.year, 12, 31)




# -----------------------
# Excel import function
# -----------------------




VALID_REFILL_MONTHS = [0.5, 1, 2, 2.8, 3, 4, 5, 6]

def import_refills_from_excel(file):

    MAX_FILE_SIZE =  5 * 1024 * 1024 * 1024  # 5 GB

    if file.size > MAX_FILE_SIZE:
        raise ValidationError("File size exceeds 50 MB limit")


    file.seek(0)

    df = pd.read_excel(file)

    df.columns = df.columns.str.strip().str.replace('\n','').str.replace('\r','').str.lower()

    required_columns_map = {
        'unique id': 'unique id',
        'last pickup date (yyyy-mm-dd)': 'last pickup date (yyyy-mm-dd)',
        'months of arv refill': 'months of arv refill',
        'current art regimen': 'current art regimen',
        'case manager': 'case manager',
        'sex': 'sex',
        'current art status': 'current art status',
        'facility name': 'facility name',

        'art start date (yyyy-mm-dd)': 'art start date (yyyy-mm-dd)',
        'date of viral load sample collection (yyyy-mm-dd)': 'date of viral load sample collection (yyyy-mm-dd)',
        'current viral load (c/ml)': 'current viral load (c/ml)',

        'date of tpt start (yyyy-mm-dd)': 'date of tpt start (yyyy-mm-dd)',
        'tpt completion date (yyyy-mm-dd)': 'tpt completion date (yyyy-mm-dd)',

        'date of commencement of eac (yyyy-mm-dd)': 'date of commencement of eac (yyyy-mm-dd)',
        'number of eac sessions completed': 'number of eac sessions completed',

        # NEW TB COLUMNS
        'age': 'age',
        'date of tb screening (yyyy-mm-dd)': 'date of tb screening (yyyy-mm-dd)',
        'tb screening type': 'tb screening type',
        'tb status': 'tb status',
        'date of tb sample collection (yyyy-mm-dd)': 'date of tb sample collection (yyyy-mm-dd)',
        'date of tb diagnostic result received (yyyy-mm-dd)': 'date of tb diagnostic result received (yyyy-mm-dd)',
        'tb diagnostic result': 'tb diagnostic result',
    }

    missing_columns = [col for col in required_columns_map.values() if col not in df.columns]

    if missing_columns:
        raise ValidationError(f"Missing column(s): {', '.join(missing_columns)}")

    df = df[df['current art status'].isin(['Active', 'Active Restart'])]

    if df.empty:
        raise ValidationError("No Active or Active Restart patients found.")

    df['facility name'] = df['facility name'].astype(str).str.strip()

    facilities = {
        f.name.strip(): f for f in Facility.objects.filter(
            name__in=df['facility name'].unique()
        )
    }

    missing_facilities = set(df['facility name'].unique()) - set(facilities.keys())

    if missing_facilities:
        raise ValidationError(
            f"Facilities not in system: {', '.join(missing_facilities)}"
        )

    validated_rows = []

    for _, row in df.iterrows():

        unique_id = row['unique id']

        try:
            last_pickup_date = pd.to_datetime(
                row['last pickup date (yyyy-mm-dd)']
            ).date()
        except Exception:
            raise ValidationError(f"Invalid Last Pickup Date for {unique_id}")

        try:
            months = float(row['months of arv refill'])
        except Exception:
            raise ValidationError(
                f"Invalid Months of ARV Refill for {unique_id}"
            )

        if months not in VALID_REFILL_MONTHS:
            raise ValidationError(
                f"Invalid refill months {months} for {unique_id}"
            )

        facility_obj = facilities[row['facility name']]

        next_appointment = last_pickup_date + timedelta(days=months * 30)

        # ================= OPTIONAL FIELDS =================

        art_start_date = pd.to_datetime(
            row['art start date (yyyy-mm-dd)']
        ).date() if pd.notnull(row['art start date (yyyy-mm-dd)']) else None

        vl_sample_collection_date = pd.to_datetime(
            row['date of viral load sample collection (yyyy-mm-dd)']
        ).date() if pd.notnull(row['date of viral load sample collection (yyyy-mm-dd)']) else None

        vl_result = int(row['current viral load (c/ml)']) if pd.notnull(
            row['current viral load (c/ml)']) else None

        tpt_start_date = pd.to_datetime(
            row['date of tpt start (yyyy-mm-dd)']
        ).date() if pd.notnull(row['date of tpt start (yyyy-mm-dd)']) else None

        tpt_completion_date = pd.to_datetime(
            row['tpt completion date (yyyy-mm-dd)']
        ).date() if pd.notnull(row['tpt completion date (yyyy-mm-dd)']) else None

        tpt_expected_completion = (
            tpt_start_date + timedelta(days=180)
            if tpt_start_date else None
        )

        eac_start_date = pd.to_datetime(
            row['date of commencement of eac (yyyy-mm-dd)']
        ).date() if pd.notnull(row['date of commencement of eac (yyyy-mm-dd)']) else None

        eac_sessions_completed = int(
            row['number of eac sessions completed']
        ) if pd.notnull(row['number of eac sessions completed']) else 0

        # ================= NEW TB FIELDS =================

        age = int(row['age']) if pd.notnull(row['age']) else None

        tb_screening_date = pd.to_datetime(
            row['date of tb screening (yyyy-mm-dd)']
        ).date() if pd.notnull(row['date of tb screening (yyyy-mm-dd)']) else None

        tb_screening_type = str(row['tb screening type']).strip() if pd.notnull(
            row['tb screening type']) else None

        tb_status = str(row['tb status']).strip() if pd.notnull(
            row['tb status']) else None

        tb_sample_collection_date = pd.to_datetime(
            row['date of tb sample collection (yyyy-mm-dd)']
        ).date() if pd.notnull(
            row['date of tb sample collection (yyyy-mm-dd)']
        ) else None

        tb_result_received_date = pd.to_datetime(
            row['date of tb diagnostic result received (yyyy-mm-dd)']
        ).date() if pd.notnull(
            row['date of tb diagnostic result received (yyyy-mm-dd)']
        ) else None

        tb_diagnostic_result = str(row['tb diagnostic result']).strip() if pd.notnull(
            row['tb diagnostic result']) else None

        validated_rows.append(

            Refill(

                facility=facility_obj,
                unique_id=unique_id,

                last_pickup_date=last_pickup_date,
                months_of_refill_days=months,
                next_appointment=next_appointment,

                current_regimen=str(row['current art regimen']).strip(),
                case_manager=str(row['case manager']).strip(),
                sex=str(row['sex']).strip(),

                current_art_status=row['current art status'].strip(),

                # VL
                art_start_date=art_start_date,
                vl_sample_collection_date=vl_sample_collection_date,
                vl_result=vl_result,

                # TPT
                tpt_start_date=tpt_start_date,
                tpt_completion_date=tpt_completion_date,
                tpt_expected_completion=tpt_expected_completion,

                # EAC
                eac_start_date=eac_start_date,
                eac_sessions_completed=eac_sessions_completed,

                # NEW TB
                age=age,
                tb_screening_date=tb_screening_date,
                tb_screening_type=tb_screening_type,
                tb_status=tb_status,
                tb_sample_collection_date=tb_sample_collection_date,
                tb_result_received_date=tb_result_received_date,
                tb_diagnostic_result=tb_diagnostic_result,
            )

        )

    facility_ids = {obj.facility.id for obj in validated_rows}

    with transaction.atomic():

        for facility_id in facility_ids:
            Refill.objects.filter(facility_id=facility_id).delete()

        Refill.objects.bulk_create(validated_rows, batch_size=1000)

    return len(validated_rows)



def upload_excel(request):
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)
        if not request.FILES:
            messages.error(request, "No file was uploaded.")
            return redirect('upload_excel')
        if form.is_valid():
            excel_file = form.cleaned_data['file']
            if excel_file.size >  50 * 1024 * 1024 :
                messages.error(request, "File size exceeds 50 GB limit.")
                return redirect('upload_excel')
            try:
                count = import_refills_from_excel(excel_file)
                messages.success(request, f"Excel uploaded! {count} records imported.")
                return redirect('upload_excel')
            except ValidationError as e:
                messages.error(request, str(e))
                return redirect('upload_excel')
            except Exception as e:
                messages.error(request, f"Upload failed: {str(e)}")
                return redirect('upload_excel')
        else:
            messages.error(request, "Form validation failed.")
            return redirect('upload_excel')
    else:
        form = UploadExcelForm()
    return render(request, 'upload.html', {'form': form})
# ----------------------------
# EXCEL UPLOAD VIEW
# ----------------------------


def upload_excel(request):
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)
        if not request.FILES:
            messages.error(request, "No file was uploaded.")
            return redirect('upload_excel')

        if form.is_valid():
            excel_file = form.cleaned_data['file']
            if excel_file.size > 1073741824:
                messages.error(request, "File size exceeds the 1GB limit.")
                return redirect('upload_excel')
            try:
                count = import_refills_from_excel(excel_file)
                messages.success(request, f"Excel uploaded successfully! {count} records imported.")
                return redirect('upload_excel')
            except ValidationError as e:
                messages.error(request, str(e))
                return redirect('upload_excel')
            except Exception as e:
                messages.error(request, f"Upload failed: {str(e)}")
                return redirect('upload_excel')
        else:
            messages.error(request, "Form validation failed.")
            return redirect('upload_excel')
    else:
        form = UploadExcelForm()

    return render(request, 'upload.html', {'form': form})


@login_required
def dashboard(request):
    today = timezone.now().date()

    # ================= QUARTER LOGIC =================
    if today.month in [1, 2, 3]:
        quarter_start = date(today.year, 1, 1)
        quarter_end = date(today.year, 3, 31)
    elif today.month in [4, 5, 6]:
        quarter_start = date(today.year, 4, 1)
        quarter_end = date(today.year, 6, 30)
    elif today.month in [7, 8, 9]:
        quarter_start = date(today.year, 7, 1)
        quarter_end = date(today.year, 9, 30)
    else:
        quarter_start = date(today.year, 10, 1)
        quarter_end = date(today.year, 12, 31)

    week_end = today + timedelta(days=7)
    month_start = today.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()

    # ================= TX_CURR (BASE POPULATION) =================
    refills = Refill.objects.filter(
        current_art_status__in=["Active", "Active Restart", "Restart"]
    )

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    # ================= COUNTERS =================
    daily_expected = 0
    daily_refills = 0
    weekly_expected = 0
    monthly_expected = 0
    monthly_missed_total = 0
    iit_total = 0

    eac_count = 0
    post_eac_vl_count = 0
    ahd_count = 0

    # ================= VL POPULATIONS =================
    vl_population = []
    quarterly_population = []

    # ================= VL METRICS =================
    overall_vl_sample_collected = 0
    overall_vl_result_received = 0

    quarterly_vl_sample_collected = 0
    quarterly_vl_result_received = 0

    suppressed_count = 0

    # ================= TB =================
    tb_screened = 0
    tb_presumptive = 0
    tb_sample_collected = 0
    tb_result_received = 0
    tb_positive = 0
    tb_negative = 0

    # ================= MAIN LOOP =================
    for r in refills:

        # ---------------- AHD / EAC ----------------
        if r.ahd:
            ahd_count += 1

        if r.eac:
            eac_count += 1

            if (
                r.eac_sessions_completed >= 3 and
                r.vl_result is not None and
                r.vl_result >= 1000
            ):
                post_eac_vl_count += 1

        # ---------------- APPOINTMENTS ----------------
        if r.next_appointment and month_start <= r.next_appointment <= month_end:
            monthly_expected += 1

            if r.next_appointment == today:
                daily_expected += 1

            if today <= r.next_appointment <= week_end:
                weekly_expected += 1

        if r.last_pickup_date == today:
            daily_refills += 1

        days_missed = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        if month_start <= r.next_appointment <= month_end and days_missed > 0:
            monthly_missed_total += 1

        if month_start <= r.next_appointment <= month_end and days_missed >= 28:
            iit_total += 1

        # ================= VL DENOMINATOR (TX_CURR STANDARD) =================
        if (
            r.art_start_date and
            (today - r.art_start_date).days >= 180 and
            r.patient_discontinued != "Y"
        ):
            vl_population.append(r)

            # Quarterly VL eligibility (time-based, NOT appointment-based)
            if r.vl_sample_collection_date:
                next_due = r.vl_sample_collection_date + relativedelta(
                    months=r.vl_interval_months
                )

                if quarter_start <= next_due <= quarter_end:
                    quarterly_population.append(r)

        # ================= VL SAMPLE COLLECTION =================
        if r.vl_sample_collection_date:
            overall_vl_sample_collected += 1

            if quarter_start <= r.vl_sample_collection_date <= quarter_end:
                quarterly_vl_sample_collected += 1

        # ================= VL RESULT =================
        if r.vl_result is not None:
            overall_vl_result_received += 1

            if quarter_start <= r.vl_sample_collection_date <= quarter_end:
                quarterly_vl_result_received += 1

            if r.vl_result < 1000:
                suppressed_count += 1

        # ================= TB =================
        if r.tb_screening_date and month_start <= r.tb_screening_date <= month_end:
            tb_screened += 1

        if r.tb_status == "Presumptive TB":
            tb_presumptive += 1

        if r.tb_sample_collection_date and month_start <= r.tb_sample_collection_date <= month_end:
            tb_sample_collected += 1

        if r.tb_result_received_date and month_start <= r.tb_result_received_date <= month_end:
            tb_result_received += 1

        if r.tb_diagnostic_result == "Positive":
            tb_positive += 1
        elif r.tb_diagnostic_result == "Negative":
            tb_negative += 1

    # ================= FINAL VL METRICS =================
    vl_denominator = len(vl_population)
    quarterly_vl_denominator = len(quarterly_population)

    vl_coverage = (
        round((quarterly_vl_sample_collected / quarterly_vl_denominator) * 100, 1)
        if quarterly_vl_denominator else 0
    )

    vl_suppression_rate = (
        round((suppressed_count / overall_vl_result_received) * 100, 1)
        if overall_vl_result_received else 0
    )

    vl_coverage_gap = max(0, quarterly_vl_denominator - quarterly_vl_sample_collected)

    return render(request, "dashboard.html", {
        "facilities": facilities,
        "selected_facility": facility_id,

        # ================= APPOINTMENTS =================
        "daily_expected": daily_expected,
        "daily_refills": daily_refills,
        "weekly_expected": weekly_expected,
        "monthly_expected": monthly_expected,
        "monthly_missed_total": monthly_missed_total,
        "iit_total": iit_total,

        # ================= PROGRAM =================
        "eac_count": eac_count,
        "post_eac_vl_count": post_eac_vl_count,
        "ahd_count": ahd_count,

        # ================= VL =================
        "vl_denominator": vl_denominator,
        "quarterly_vl_denominator": quarterly_vl_denominator,

        "quarter_vl_collected": quarterly_vl_sample_collected,
        "quarter_vl_result_received": quarterly_vl_result_received,

        "vl_coverage": vl_coverage,
        "vl_coverage_gap": vl_coverage_gap,
        "vl_suppression_rate": vl_suppression_rate,

        # ================= TB =================
        "tb_screened": tb_screened,
        "tb_presumptive": tb_presumptive,
        "tb_sample_collected": tb_sample_collected,
        "tb_result_received": tb_result_received,
        "tb_positive": tb_positive,
        "tb_negative": tb_negative,
    })
# ================================

@login_required
def refill_list(request):

    today = timezone.now().date()
    week_end = today + timedelta(days=7)

    # ================= QUARTER LOGIC =================
    if today.month in [1, 2, 3]:
        quarter_start = date(today.year, 1, 1)
        quarter_end = date(today.year, 3, 31)
    elif today.month in [4, 5, 6]:
        quarter_start = date(today.year, 4, 1)
        quarter_end = date(today.year, 6, 30)
    elif today.month in [7, 8, 9]:
        quarter_start = date(today.year, 7, 1)
        quarter_end = date(today.year, 9, 30)
    else:
        quarter_start = date(today.year, 10, 1)
        quarter_end = date(today.year, 12, 31)

    # ================= FILTERS =================
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    search_unique_id = request.GET.get("search_unique_id")

    facilities = Facility.objects.all()
    case_managers = Refill.objects.values_list("case_manager", flat=True).distinct()

    refills = Refill.objects.all()

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    if start_date:
        refills = refills.filter(next_appointment__gte=start_date)

    if end_date:
        refills = refills.filter(next_appointment__lte=end_date)

    if search_unique_id:
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    # ================= PROCESS PATIENTS =================
    for r in refills:

        # ---------------- DAYS MISSED ----------------
        r.days_missed_display = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        r.missed_appointment = r.days_missed_display > 0

        # =====================================================
        # ✅ VL STATUS (TRUTH SOURCE = MODEL PROPERTY)
        # =====================================================
        r.vl_status = "Eligible" if r.is_vl_eligible_program else "Not Eligible"

        # =====================================================
        # QUARTER VL STATUS (PROGRAM ACCURATE)
        # =====================================================
        if r.is_vl_eligible_program:

            # never had VL
            if not r.vl_sample_collection_date:
                r.vl_quarter_status = "Due (No VL Ever)"

            # sample outside quarter
            elif r.vl_sample_collection_date < quarter_start:
                r.vl_quarter_status = "Due (Overdue)"

            # sample inside quarter
            elif quarter_start <= r.vl_sample_collection_date <= quarter_end:
                r.vl_quarter_status = "Done This Quarter"

            else:
                r.vl_quarter_status = "Pending"
        else:
            r.vl_quarter_status = "Not Eligible"

    # ================= PERIOD FILTERS =================
    daily_expected = refills.filter(next_appointment=today)

    weekly_expected = refills.filter(next_appointment__range=[today, week_end])

    monthly_expected = refills.filter(
        next_appointment__year=today.year,
        next_appointment__month=today.month
    )

    periods = [
        {
            "name": "Daily",
            "page_obj": Paginator(daily_expected, 10).get_page(request.GET.get("daily_page"))
        },
        {
            "name": "Weekly",
            "page_obj": Paginator(weekly_expected, 10).get_page(request.GET.get("weekly_page"))
        },
        {
            "name": "Monthly",
            "page_obj": Paginator(monthly_expected, 10).get_page(request.GET.get("monthly_page"))
        },
    ]

    return render(request, "refill_list.html", {
        "facilities": facilities,
        "case_managers": case_managers,
        "selected_facility": facility_id,
        "selected_case_manager": selected_case_manager,
        "periods": periods,
        "today": today,
        "search_unique_id": search_unique_id,
        "query_params": request.GET.urlencode(),
    })
    
@login_required
def export_refills_view(request):

    today = timezone.now().date()

    # ================= FILTERS (SAME AS VIEW) =================
    refills = Refill.objects.all()

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    search_unique_id = request.GET.get("search_unique_id")

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    if start_date:
        refills = refills.filter(next_appointment__gte=start_date)

    if end_date:
        refills = refills.filter(next_appointment__lte=end_date)

    if search_unique_id:
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    # ================= EXPORT =================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expected Refills Data"

    headers = [
        'Unique ID',
        'Facility',
        'Sex',
        'Current Regimen',
        'Case Manager',
        'Last Pickup Date',
        'Next Appointment',
        'Days Missed',
        'VL Result (c/ml)',
        'VL Eligibility',
        'EAC Status',
        'AHD Status',
        'TPT Start Date',
        'TPT Completion Date',
        'TPT Status',
        'TB Cascade Status',
        'TB Screening Date',
        'TB Result Received Date',
        'TB Diagnostic Result',
        'Remark',
        'Tracking Date 1',
        'Tracking Date 2',
        'Tracking Date 3',
        'Tracked By',
        'Patient Discontinued',
        'Discontinued Reason',
        'Discontinued Date',
        'Returned Date',
    ]

    ws.append(headers)

    for refill in refills:

        next_appointment = refill.next_appointment

        # SAME LOGIC AS VIEW
        days_missed = (
            (today - next_appointment).days
            if next_appointment and next_appointment < today
            else 0
        )

        # SAME VL SOURCE AS VIEW
        vl_eligibility = "Eligible" if refill.is_vl_eligible_program else "Not Eligible"

        ahd_status = "Eligible" if refill.current_art_status == "Restart" else "Not Eligible"

        ws.append([
            refill.unique_id,
            refill.facility.name if refill.facility else "",
            refill.sex,
            refill.current_regimen,
            refill.case_manager,
            refill.last_pickup_date.strftime("%Y-%m-%d") if refill.last_pickup_date else "",
            next_appointment.strftime("%Y-%m-%d") if next_appointment else "",
            days_missed,
            refill.vl_result or "",
            vl_eligibility,
            refill.eac_status,
            ahd_status,
            refill.tpt_start_date.strftime("%Y-%m-%d") if refill.tpt_start_date else "",
            refill.tpt_completion_date.strftime("%Y-%m-%d") if refill.tpt_completion_date else "",
            refill.tpt_status,
            refill.tb_cascade_status or "",
            refill.tb_screening_date.strftime("%Y-%m-%d") if refill.tb_screening_date else "",
            refill.tb_result_received_date.strftime("%Y-%m-%d") if refill.tb_result_received_date else "",
            refill.tb_diagnostic_result or "",
            refill.remark or "",
            refill.tracking_date_1.strftime("%Y-%m-%d") if refill.tracking_date_1 else "",
            refill.tracking_date_2.strftime("%Y-%m-%d") if refill.tracking_date_2 else "",
            refill.tracking_date_3.strftime("%Y-%m-%d") if refill.tracking_date_3 else "",
            refill.tracked_by or "",
            dict(refill._meta.get_field('patient_discontinued').choices).get(refill.patient_discontinued, "") if refill.patient_discontinued else "",
            refill.discontinued_reason or "",
            refill.discontinued_date.strftime("%Y-%m-%d") if refill.discontinued_date else "",
            refill.returned_date.strftime("%Y-%m-%d") if refill.returned_date else "",
        ])

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    response["Content-Disposition"] = f'attachment; filename="Expected_Refills_{today}.xlsx"'

    wb.save(response)

    return response


# ================= REFILL CREATE =================




@login_required
def refill_create_or_update(request, pk=None):
    today = timezone.now().date()
    refill = get_object_or_404(Refill, pk=pk) if pk else Refill()

    if request.method == "POST":
        form = RefillForm(request.POST, instance=refill)
        if form.is_valid():
            refill = form.save(commit=False)

            # ------------------ Next Appointment ------------------
            if refill.last_pickup_date and refill.months_of_refill_days:
                refill.next_appointment = refill.last_pickup_date + timedelta(days=float(refill.months_of_refill_days)*30)
            else:
                refill.next_appointment = None

            # ------------------ Missed Appointment ------------------
            refill.missed_appointment = refill.next_appointment < today if refill.next_appointment else False

            # ------------------ TPT expected completion ------------------
            if refill.tpt_start_date:
                refill.tpt_expected_completion = refill.tpt_start_date + timedelta(days=180)

            # ------------------ Save ------------------
            refill.save()
            return redirect("refill_list")
    else:
        form = RefillForm(instance=refill)

    return render(request, "refill_form.html", {"form": form, "today": today})







def refill_add_or_update(request, unique_id=None):
    """
    Handles adding a new refill or updating an existing record.
    Allows updating tracking, remarks, or discontinuation even if
    last pickup / refill is empty.
    """
    # Try to get existing Refill for this unique_id
    if unique_id:
        refill = Refill.objects.filter(unique_id=unique_id).order_by('-last_pickup_date').first()
    else:
        refill = None

    if request.method == 'POST':
        form = RefillForm(request.POST, instance=refill)
        if form.is_valid():
            refill_instance = form.save(commit=False)

            # If next_appointment is empty but last_pickup_date exists, auto-calc next appointment
            if refill_instance.last_pickup_date and not refill_instance.next_appointment:
                refill_instance.next_appointment = refill_instance.last_pickup_date + timezone.timedelta(
                    days=(refill_instance.months_of_refill_days or 0) * 30
                )

            refill_instance.save()
            messages.success(request, 'Refill record saved successfully.')
            return redirect('refill_list')
        else:
            messages.error(request, 'Please fix the errors below.')

    else:
        # GET request
        form = RefillForm(instance=refill)

    # Latest TB info (for your form)
    latest_tb = getattr(refill, 'latest_tb', None)

    context = {
        'form': form,
        'latest_tb': latest_tb,
    }
    return render(request, 'refill_form.html', context)





@login_required
def refill_create(request, unique_id=None):
    if unique_id:
        # Fetch the refill by unique_id if passed
        refill = get_object_or_404(Refill, unique_id=unique_id)
    else:
        refill = None  # New refill if no unique_id is passed

    today = timezone.now().date()  # ensure datetime.date object

    if refill:
        # Ensure next_appointment is a date
        if isinstance(refill.next_appointment, datetime):
            refill_next_appointment = refill.next_appointment.date()
        else:
            refill_next_appointment = refill.next_appointment

        if refill_next_appointment and refill_next_appointment < today:
            # Optional: any logic for past appointments
            print("This refill's next appointment is in the past.")

    if request.method == 'POST':
        if refill:
            form = RefillForm(request.POST, instance=refill)
        else:
            form = RefillForm(request.POST)

        if form.is_valid():
            form.save()  # vl_status is read-only now
            return redirect('refill_list')

    else:
        if refill:
            form = RefillForm(instance=refill)
        else:
            form = RefillForm()

    return render(request, 'refill_form.html', {'form': form})




# ================= REFILL UPDATE =================


@login_required
def refill_update(request, pk):
    """
    Update an existing refill and auto-recalculate next appointment.
    Supports single-field updates (e.g., only 'remark').
    """
    refill = get_object_or_404(Refill, pk=pk)
    form = RefillForm(request.POST or None, instance=refill)

    if form.is_valid():
        refill = form.save(commit=False)

        # Auto recalculate next appointment if last pickup or refill duration changed
        if refill.last_pickup_date and refill.months_of_refill_days:
            days = float(refill.months_of_refill_days) * 30
            refill.next_appointment = refill.last_pickup_date + timedelta(days=days)

        # Auto update TPT expected completion
        if refill.tpt_start_date:
            refill.tpt_expected_completion = refill.tpt_start_date + timedelta(days=180)

        refill.save()  # saves all changes
        return redirect('refill_list')  # or your dashboard

    context = {
        "form": form
    }
    return render(request, "refill_form.html", context)
# ================= REFILL ADD OR UPDATE =================



@login_required
def track_refills(request):

    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")

    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # ================= PARSE DATES =================
    if start_date:
        try:
            start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date = None

    if end_date:
        try:
            end_date = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            end_date = None

    facilities = Facility.objects.all()
    case_managers = Refill.objects.values_list("case_manager", flat=True).distinct()

    # ================= MONTH RANGE =================
    month_start = today.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    refills = Refill.objects.select_related("facility")

    # ================= FILTERS =================
    if start_date and end_date:
        refills = refills.filter(last_pickup_date__range=[start_date, end_date])
    elif start_date:
        refills = refills.filter(last_pickup_date__gte=start_date)
    elif end_date:
        refills = refills.filter(last_pickup_date__lte=end_date)
    else:
        refills = refills.filter(last_pickup_date__range=(month_start, month_end))

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    filtered_refills = []

    # ================= MAIN LOOP =================
    for r in refills:

        # ---------------- DAYS MISSED ----------------
        r.days_missed_display = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        # =====================================================
        # 🔥 STANDARDIZED VL ELIGIBILITY SOURCE
        # =====================================================
        vl_eligible = r.is_vl_eligible_program

        # =====================================================
        # 🧠 SUPPRESSION STATE (CLINICAL LOGIC)
        # =====================================================
        is_suppressed = (
            r.vl_result is not None and r.vl_result < 1000
        )

        # =====================================================
        # 🔬 VL STATUS (CLEAN PROGRAM LOGIC)
        # =====================================================

        if not vl_eligible:
            if r.patient_discontinued == "Y":
                r.vl_status = "Not Eligible (Discontinued)"

            elif r.days_missed_display >= 28:
                r.vl_status = "Not Eligible (IIT)"

            elif r.art_start_date and (today - r.art_start_date).days < 180:
                r.vl_status = "Not Eligible (ART < 6 months)"

            else:
                r.vl_status = "Not Eligible"

        else:
            # ================= NEVER HAD VL =================
            if not r.vl_sample_collection_date:
                r.vl_status = "Eligible (No VL Yet)"

            # ================= HAS VL =================
            else:

                if r.vl_result is None:
                    r.vl_status = "VL Pending"

                elif is_suppressed:
                    # suppressed clients
                    delta = relativedelta(today, r.vl_sample_collection_date)
                    months_since = delta.years * 12 + delta.months

                    if months_since >= 3:
                        r.vl_status = "Repeat VL Due (Suppressed)"
                    else:
                        r.vl_status = "Monitoring Suppressed"

                else:
                    # suppressed NOT
                    if vl_eligible:
                        r.vl_status = "Routine VL Due"
                    else:
                        r.vl_status = "VL Up to Date"

        # =====================================================
        # 🧾 DISPLAY CLEANING (UNCHANGED)
        # =====================================================
        r.age_display = r.age or "-"

        r.tracking_date_1_display = r.tracking_date_1
        r.tracking_date_2_display = r.tracking_date_2
        r.tracking_date_3_display = r.tracking_date_3

        r.patient_discontinued_display = (
            dict(r._meta.get_field('patient_discontinued').choices).get(
                r.patient_discontinued, ""
            ) if r.patient_discontinued else "-"
        )

        filtered_refills.append(r)

    # ================= PAGINATION =================
    paginator = Paginator(filtered_refills, 10)
    page_number = request.GET.get("page")
    page_obj = paginator.get_page(page_number)

    return render(request, "track_refills.html", {
        "facilities": facilities,
        "case_managers": case_managers,
        "selected_facility": facility_id,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "page_obj": page_obj,
    })
@login_required
def export_track_refills_view(request):

    today = timezone.now().date()

    # ================= FILTERS (SAME AS VIEW) =================
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # PARSE DATES (SAME AS VIEW)
    if start_date:
        try:
            start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date = None

    if end_date:
        try:
            end_date = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            end_date = None

    month_start = today.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    refills = Refill.objects.select_related("facility")

    if start_date and end_date:
        refills = refills.filter(last_pickup_date__range=[start_date, end_date])
    elif start_date:
        refills = refills.filter(last_pickup_date__gte=start_date)
    elif end_date:
        refills = refills.filter(last_pickup_date__lte=end_date)
    else:
        refills = refills.filter(last_pickup_date__range=(month_start, month_end))

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # ================= EXPORT =================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track Refills Data"

    headers = [
        'Unique ID', 'Facility', 'Last Pickup Date', 'Refill Days', 'Sex',
        'Current Regimen', 'Case Manager', 'Next Appointment',
        'Days Missed', 'VL Status',
        'Tracking Date 1',
        'Tracking Date 2',
        'Tracking Date 3',
        'Tracked By',
        'Patient Discontinued',
        'Discontinued Reason',
        'Discontinued Date',
        'Returned Date',
    ]
    ws.append(headers)

    for r in refills:

        r.calculate_dates()

        next_appointment = r.next_appointment
        last_pickup = r.last_pickup_date

        # SAME AS VIEW
        days_missed = (
            (today - next_appointment).days
            if next_appointment and next_appointment < today
            else 0
        )

        # SAME VL LOGIC AS VIEW
        vl_eligible = r.is_vl_eligible_program

        is_suppressed = (
            r.vl_result is not None and r.vl_result < 1000
        )

        if not vl_eligible:
            if r.patient_discontinued == "Y":
                vl_status = "Not Eligible (Discontinued)"
            elif days_missed >= 28:
                vl_status = "Not Eligible (IIT)"
            elif r.art_start_date and (today - r.art_start_date).days < 180:
                vl_status = "Not Eligible (ART < 6 months)"
            else:
                vl_status = "Not Eligible"
        else:
            if not r.vl_sample_collection_date:
                vl_status = "Eligible (No VL Yet)"
            else:
                if r.vl_result is None:
                    vl_status = "VL Pending"
                elif is_suppressed:
                    delta = relativedelta(today, r.vl_sample_collection_date)
                    months_since = delta.years * 12 + delta.months

                    if months_since >= 3:
                        vl_status = "Repeat VL Due (Suppressed)"
                    else:
                        vl_status = "Monitoring Suppressed"
                else:
                    if vl_eligible:
                        vl_status = "Routine VL Due"
                    else:
                        vl_status = "VL Up to Date"

        ws.append([
            r.unique_id,
            r.facility.name if r.facility else "",
            last_pickup.strftime("%Y-%m-%d") if last_pickup else "Never Picked",
            r.months_of_refill_days,
            r.sex,
            r.current_regimen,
            r.case_manager or "",
            next_appointment.strftime("%Y-%m-%d") if next_appointment else "",
            days_missed,
            vl_status,
            r.tracking_date_1.strftime("%Y-%m-%d") if r.tracking_date_1 else "",
            r.tracking_date_2.strftime("%Y-%m-%d") if r.tracking_date_2 else "",
            r.tracking_date_3.strftime("%Y-%m-%d") if r.tracking_date_3 else "",
            r.tracked_by or "",
            dict(r._meta.get_field('patient_discontinued').choices).get(r.patient_discontinued, "") if r.patient_discontinued else "",
            r.discontinued_reason or "",
            r.discontinued_date.strftime("%Y-%m-%d") if r.discontinued_date else "",
            r.returned_date.strftime("%Y-%m-%d") if r.returned_date else "",
        ])

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    response['Content-Disposition'] = f'attachment; filename="Track_Refills_{today}.xlsx"'

    wb.save(response)
    return response



# @login_required
# def export_all_track_refills_excel(request):
#     """
#     Export all refills for tracking to Excel.
#     """
#     refills = Refill.objects.all()  # download ALL data
#     return export_track_refills_to_excel(refills)



@login_required
def daily_refill_list(request):

    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")

    facilities = Facility.objects.all()

    refills = Refill.objects.filter(next_appointment=today)

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # ================= CLEAN CASE MANAGERS =================
    case_managers = (
        Refill.objects
        .exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )

    case_managers = sorted(set(cm.strip() for cm in case_managers if cm))

    # ================= ENRICH OBJECTS =================
    for r in refills:

        # ---------------- AGE ----------------
        r.age_display = r.age if r.age is not None else "Unknown"

        # ---------------- TB ----------------
        r.tb_status_display = r.tb_status if r.tb_status else "Not Screened"

        # ---------------- MISSED DAYS ----------------
        r.days_missed_display = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        # =====================================================
        # 🔥 SINGLE SOURCE OF TRUTH (PROGRAM VL ELIGIBILITY)
        # =====================================================
        vl_eligible = r.is_vl_eligible_program

        # =====================================================
        # 🧠 SUPPRESSION LOGIC (ONLY RESULT BASED)
        # =====================================================
        is_suppressed = (
            r.vl_result is not None and r.vl_result < 1000
        )

        # =====================================================
        # 🔬 CLEAN VL STATUS (PEPFAR STYLE LOGIC)
        # =====================================================

        if not vl_eligible:

            if r.patient_discontinued == 'Y':
                r.vl_status_display = "Not Eligible (Discontinued)"

            elif r.days_missed_display >= 28:
                r.vl_status_display = "Not Eligible (IIT)"

            elif r.art_start_date and (today - r.art_start_date).days < 180:
                r.vl_status_display = "Not Eligible (ART < 6 Months)"

            else:
                r.vl_status_display = "Not Eligible"

        else:

            # ================= NO VL HISTORY =================
            if not r.vl_sample_collection_date:
                r.vl_status_display = "Eligible (VL Due)"

            # ================= HAS VL =================
            else:

                if r.vl_result is None:
                    r.vl_status_display = "VL Pending"

                elif is_suppressed:
                    r.vl_status_display = "Suppressed (<1000)"

                else:
                    r.vl_status_display = "Not Suppressed (≥1000)"

        # ================= EAC =================
        r.eac_status_display = r.eac_status

        # ================= AHD =================
        r.ahd_display = "Eligible" if r.ahd else "Not Eligible"

    return render(request, "daily_refill_list.html", {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "refills": refills,
    })
    
    
@login_required
def missed_refills(request):
    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    search_unique_id = request.GET.get("search_unique_id")

    # ================= BASE QUERYSET =================
    refills = Refill.objects.filter(
        current_art_status__in=["Active", "Active Restart", "Restart"]
    ).select_related("facility")

    # ================= FILTERS =================
    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if case_manager:
        refills = refills.filter(case_manager__iexact=case_manager.strip())

    if search_unique_id:
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    if start_date:
        try:
            refills = refills.filter(next_appointment__gte=start_date)
        except:
            pass

    if end_date:
        try:
            refills = refills.filter(next_appointment__lte=end_date)
        except:
            pass

    # ================= ONLY MISSED =================
    missed_list = refills.filter(
        next_appointment__lt=today
    ).order_by("next_appointment")

    # ================= ENRICH DATA =================
    for r in missed_list:

        # days missed (use model property if available)
        r.days_missed_display = r.days_missed

        # IIT status (already standardized in model)
        r.iit_status_display = r.iit_status

        # =====================================================
        # ✅ UNIFIED VL LOGIC (PROGRAM STANDARD)
        # =====================================================

        if not r.is_vl_eligible_program:
            r.vl_status_display = "Not Eligible"

        else:
            # ART eligible population

            if not r.vl_sample_collection_date:
                r.vl_status_display = "Eligible (No VL Yet)"

            else:
                # Age-based interval rule
                interval_months = 6 if (r.age and r.age < 15) else 12

                delta_months = (
                    (today.year - r.vl_sample_collection_date.year) * 12 +
                    (today.month - r.vl_sample_collection_date.month)
                )

                if delta_months >= interval_months:
                    r.vl_status_display = "Eligible (Due)"
                else:
                    r.vl_status_display = "Not Due"

    # ================= PAGINATION =================
    paginator = Paginator(missed_list, 25)
    page_number = request.GET.get("page")
    page_obj = paginator.get_page(page_number)

    # ================= CASE MANAGERS =================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )

    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm})

    query_params = request.GET.copy()
    query_params.pop("page", None)

    return render(request, "missed_refills.html", {
        "page_obj": page_obj,
        "today": today,
        "total_missed": missed_list.count(),

        "facilities": Facility.objects.all(),
        "case_managers": case_managers,

        "selected_facility": facility_id,
        "selected_case_manager": case_manager,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "search_unique_id": search_unique_id,

        "query_params": query_params.urlencode(),
    })
    
    
    
    
@login_required
def export_missed_refills_view(request):

    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    search_unique_id = request.GET.get("search_unique_id")

    # ================= BASE QUERYSET =================
    refills = Refill.objects.filter(
        current_art_status__in=["Active", "Active Restart", "Restart"]
    ).select_related("facility")

    # ================= FILTERS =================
    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if case_manager:
        refills = refills.filter(case_manager__iexact=case_manager.strip())

    if search_unique_id:
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    if start_date:
        try:
            refills = refills.filter(next_appointment__gte=start_date)
        except:
            pass

    if end_date:
        try:
            refills = refills.filter(next_appointment__lte=end_date)
        except:
            pass

    # ================= ONLY MISSED =================
    missed_list = refills.filter(
        next_appointment__lt=today
    ).order_by("next_appointment")

    # ================= EXPORT =================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Missed Refills Data"

    headers = [
        'Unique ID',
        'Facility',
        'Sex',
        'Current Regimen',
        'Case Manager',
        'Last Pickup Date',
        'Next Appointment',
        'Days Missed',
        'IIT Status',
        'VL Status',
        'Tracking Date 1',
        'Tracking Date 2',
        'Tracking Date 3',
        'Tracked By',
        'Patient Discontinued',
        'Discontinued Reason',
        'Discontinued Date',
        'Returned Date',
    ]
    ws.append(headers)

    for r in missed_list:

        # SAME AS VIEW
        days_missed = r.days_missed
        iit_status = r.iit_status

        # SAME VL LOGIC AS VIEW
        if not r.is_vl_eligible_program:
            vl_status = "Not Eligible"
        else:
            if not r.vl_sample_collection_date:
                vl_status = "Eligible (No VL Yet)"
            else:
                interval_months = 6 if (r.age and r.age < 15) else 12

                delta_months = (
                    (today.year - r.vl_sample_collection_date.year) * 12 +
                    (today.month - r.vl_sample_collection_date.month)
                )

                if delta_months >= interval_months:
                    vl_status = "Eligible (Due)"
                else:
                    vl_status = "Not Due"

        ws.append([
            r.unique_id,
            r.facility.name if r.facility else "",
            r.sex,
            r.current_regimen,
            r.case_manager or "",
            r.last_pickup_date.strftime("%Y-%m-%d") if r.last_pickup_date else "",
            r.next_appointment.strftime("%Y-%m-%d") if r.next_appointment else "",
            days_missed,
            iit_status,
            vl_status,
            r.tracking_date_1.strftime("%Y-%m-%d") if r.tracking_date_1 else "",
            r.tracking_date_2.strftime("%Y-%m-%d") if r.tracking_date_2 else "",
            r.tracking_date_3.strftime("%Y-%m-%d") if r.tracking_date_3 else "",
            r.tracked_by or "",
            dict(r._meta.get_field('patient_discontinued').choices).get(r.patient_discontinued, "") if r.patient_discontinued else "",
            r.discontinued_reason or "",
            r.discontinued_date.strftime("%Y-%m-%d") if r.discontinued_date else "",
            r.returned_date.strftime("%Y-%m-%d") if r.returned_date else "",
        ])

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    response['Content-Disposition'] = f'attachment; filename="Missed_Refills_{today}.xlsx"'

    wb.save(response)
    return response 
    
    
    
    
    
    
    
@login_required
def track_vl(request):
    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")

    facilities = Facility.objects.all()

    # ================= BASE QUERY =================
    refills = Refill.objects.all()

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # ================= ORDERING =================
    refills = refills.order_by("-vl_sample_collection_date")

    # ================= PROCESS EACH PATIENT =================
    for r in refills:

        # ---------------- DAYS MISSED ----------------
        r.days_missed_display = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        # =====================================================
        # ✅ NEW UNIFIED VL LOGIC (MODEL-DRIVEN)
        # =====================================================

        # 1. NOT ELIGIBLE BASED ON PROGRAM RULES
        if not r.is_vl_eligible_program:
            if r.patient_discontinued == 'Y':
                r.vl_status = "Not Eligible (Discontinued)"
            elif r.days_missed_display >= 28:
                r.vl_status = "Not Eligible (IIT)"
            elif not r.art_start_date:
                r.vl_status = "No ART"
            elif (today - r.art_start_date).days < 180:
                r.vl_status = "Not Eligible (ART < 6 months)"
            else:
                r.vl_status = "Not Eligible"

        # 2. ELIGIBLE CLIENTS
        else:

            # ================= NO VL YET =================
            if not r.vl_sample_collection_date:
                r.vl_status = "Eligible (VL Due)"

            else:

                # ================= SUPPRESSION CHECK =================
                is_suppressed = r.is_suppressed

                # ================= MONTH DIFFERENCE =================
                delta_months = (
                    (today.year - r.vl_sample_collection_date.year) * 12
                    + (today.month - r.vl_sample_collection_date.month)
                )

                # ================= UNSUPPRESSED =================
                if is_suppressed is False:

                    # WHO STANDARD: repeat VL after 3 months
                    if delta_months >= 3:
                        r.vl_status = "Repeat VL Due"
                    else:
                        r.vl_status = "Awaiting Repeat VL"

                # ================= SUPPRESSED =================
                elif is_suppressed is True:

                    interval = 12 if (r.age and r.age > 15) else 6

                    if delta_months >= interval:
                        r.vl_status = f"Eligible ({interval}-Month VL Due)"
                    else:
                        r.vl_status = "VL Up to Date"

                else:
                    r.vl_status = "VL Result Pending"

    # ================= EXPORT =================
   # REMOVE THIS BLOCK


    # ================= PAGINATION =================
    paginator = Paginator(refills, 10)
    page_number = request.GET.get("page")
    vl_refills = paginator.get_page(page_number)

    return render(request, "track_vl.html", {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": Refill.objects.values_list("case_manager", flat=True).distinct(),
        "vl_refills": vl_refills,
        "today": today,
    })
    
    
    
@login_required
def export_vl_view(request):

    today = timezone.now().date()

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")

    # ================= SAME QUERY AS VIEW =================
    refills = Refill.objects.all()

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    refills = refills.order_by("-vl_sample_collection_date")

    # ================= EXPORT =================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track VL"

    headers = [
        "Unique ID",
        "Facility",
        "VL Sample Collection Date",
        "Last Pickup Date",
        "Case Manager",
        "Days Missed",
        "VL Status",
    ]
    ws.append(headers)

    for r in refills:

        # SAME AS VIEW
        days_missed = (
            (today - r.next_appointment).days
            if r.next_appointment and r.next_appointment < today
            else 0
        )

        # SAME VL LOGIC AS VIEW
        if not r.is_vl_eligible_program:
            if r.patient_discontinued == 'Y':
                vl_status = "Not Eligible (Discontinued)"
            elif days_missed >= 28:
                vl_status = "Not Eligible (IIT)"
            elif not r.art_start_date:
                vl_status = "No ART"
            elif (today - r.art_start_date).days < 180:
                vl_status = "Not Eligible (ART < 6 months)"
            else:
                vl_status = "Not Eligible"
        else:
            if not r.vl_sample_collection_date:
                vl_status = "Eligible (VL Due)"
            else:
                is_suppressed = r.is_suppressed

                delta_months = (
                    (today.year - r.vl_sample_collection_date.year) * 12
                    + (today.month - r.vl_sample_collection_date.month)
                )

                if is_suppressed is False:
                    if delta_months >= 3:
                        vl_status = "Repeat VL Due"
                    else:
                        vl_status = "Awaiting Repeat VL"

                elif is_suppressed is True:
                    interval = 12 if (r.age and r.age > 15) else 6

                    if delta_months >= interval:
                        vl_status = f"Eligible ({interval}-Month VL Due)"
                    else:
                        vl_status = "VL Up to Date"

                else:
                    vl_status = "VL Result Pending"

        ws.append([
            r.unique_id,
            r.facility.name if r.facility else "",
            r.vl_sample_collection_date.strftime("%Y-%m-%d") if r.vl_sample_collection_date else "",
            r.last_pickup_date.strftime("%Y-%m-%d") if r.last_pickup_date else "",
            r.case_manager or "",
            days_missed,
            vl_status,
        ])

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    response["Content-Disposition"] = f'attachment; filename="Track_VL_{today}.xlsx"'

    wb.save(response)
    return response