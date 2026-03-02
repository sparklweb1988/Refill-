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















# -----------------------
# Excel import function
# -----------------------




VALID_REFILL_MONTHS = [0.5, 1, 2, 2.8, 3, 4, 5, 6]

def import_refills_from_excel(file):
    if file.size > 1073741824:
        raise ValidationError("File size exceeds 1GB limit")
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
        'number of eac sessions completed': 'number of eac sessions completed'
    }

    missing_columns = [col for col in required_columns_map.values() if col not in df.columns]
    if missing_columns:
        raise ValidationError(f"Missing column(s): {', '.join(missing_columns)}")

    df = df[df['current art status'].isin(['Active', 'Active Restart'])]
    if df.empty:
        raise ValidationError("No Active or Active Restart patients found.")

    df['facility name'] = df['facility name'].astype(str).str.strip()
    facilities = {f.name.strip(): f for f in Facility.objects.filter(name__in=df['facility name'].unique())}
    missing_facilities = set(df['facility name'].unique()) - set(facilities.keys())
    if missing_facilities:
        raise ValidationError(f"Facilities not in system: {', '.join(missing_facilities)}")

    validated_rows = []
    for _, row in df.iterrows():
        unique_id = row['unique id']
        try:
            last_pickup_date = pd.to_datetime(row['last pickup date (yyyy-mm-dd)']).date()
        except Exception:
            raise ValidationError(f"Invalid Last Pickup Date for {unique_id}")
        try:
            months = float(row['months of arv refill'])
        except Exception:
            raise ValidationError(f"Invalid Months of ARV Refill for {unique_id}")
        if months not in VALID_REFILL_MONTHS:
            raise ValidationError(f"Invalid refill months {months} for {unique_id}")
        facility_obj = facilities[row['facility name']]
        next_appointment = last_pickup_date + timedelta(days=months*30)

        # Optional fields
        art_start_date = pd.to_datetime(row['art start date (yyyy-mm-dd)']).date() if pd.notnull(row['art start date (yyyy-mm-dd)']) else None
        vl_sample_collection_date = pd.to_datetime(row['date of viral load sample collection (yyyy-mm-dd)']).date() if pd.notnull(row['date of viral load sample collection (yyyy-mm-dd)']) else None
        vl_result = int(row['current viral load (c/ml)']) if pd.notnull(row['current viral load (c/ml)']) else None
        tpt_start_date = pd.to_datetime(row['date of tpt start (yyyy-mm-dd)']).date() if pd.notnull(row['date of tpt start (yyyy-mm-dd)']) else None
        tpt_completion_date = pd.to_datetime(row['tpt completion date (yyyy-mm-dd)']).date() if pd.notnull(row['tpt completion date (yyyy-mm-dd)']) else None
        tpt_expected_completion = tpt_start_date + timedelta(days=180) if tpt_start_date else None
        eac_start_date = pd.to_datetime(row['date of commencement of eac (yyyy-mm-dd)']).date() if pd.notnull(row['date of commencement of eac (yyyy-mm-dd)']) else None
        eac_sessions_completed = int(row['number of eac sessions completed']) if pd.notnull(row['number of eac sessions completed']) else 0

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

        # EAC (ONLY REAL FIELDS)
        eac_start_date=eac_start_date,
        eac_sessions_completed=eac_sessions_completed,
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
            if excel_file.size > 1073741824:
                messages.error(request, "File size exceeds 1GB limit.")
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





# DASHBOARD

@login_required
def dashboard(request):
    today = timezone.now().date()

    # ---- DATE WINDOWS ----
    week_end = today + timedelta(days=7)
    month_start = today.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()

    refills = Refill.objects.filter(
        current_art_status__in=["Active", "Active Restart", "Restart"]
    )

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    # ---------------- COUNTERS ----------------
    daily_expected = daily_refills = weekly_expected = monthly_expected = 0
    monthly_missed_total = iit_total = 0

    eac_count = post_eac_vl_count = ahd_count = 0

    # VL METRICS (CURRENT MONTH ONLY)
    eligible_clients = []
    vl_sample_collected = 0
    vl_result_received = 0
    suppressed_count = 0

    for r in refills:

        # ---------------- AHD / EAC ----------------
        if r.ahd:
            ahd_count += 1

        if r.eac:
            eac_count += 1
            if r.post_eac_vl_due:
                post_eac_vl_count += 1

        # ---------------- EXPECTED (CURRENT MONTH ONLY) ----------------
        if r.next_appointment and month_start <= r.next_appointment <= month_end:
            monthly_expected += 1

            if r.next_appointment == today:
                daily_expected += 1
                if r.last_pickup_date == today:
                    daily_refills += 1

            if today <= r.next_appointment <= week_end:
                weekly_expected += 1

        # ---------------- MISSED / IIT (CURRENT MONTH ONLY) ----------------
        if r.next_appointment and month_start <= r.next_appointment <= month_end:
            days_missed = (today - r.next_appointment).days if r.next_appointment < today else 0

            if days_missed > 0:
                monthly_missed_total += 1

                if days_missed >= 28:
                    iit_total += 1

        # ---------------- VL METRICS (RESULTS FROM CURRENT MONTH ONLY) ----------------
        if r.is_vl_eligible:
            eligible_clients.append(r)

            if r.vl_sample_collection_date and month_start <= r.vl_sample_collection_date <= month_end:
                vl_sample_collected += 1

            if r.vl_result is not None and r.vl_sample_collection_date and month_start <= r.vl_sample_collection_date <= month_end:
                vl_result_received += 1

                if r.vl_result < 1000:
                    suppressed_count += 1

    # ---------------- FINAL CALCULATIONS ----------------
    vl_denominator = len(eligible_clients)

    vl_coverage = (
        round((vl_sample_collected / vl_denominator) * 100, 1)
        if vl_denominator else 0
    )

    vl_suppression_rate = (
        round((suppressed_count / vl_result_received) * 100, 1)
        if vl_result_received else 0
    )

    vl_coverage_gap = vl_denominator - vl_sample_collected

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "today": today,

        # Expected / IIT
        "daily_expected": daily_expected,
        "daily_refills": daily_refills,
        "weekly_expected": weekly_expected,
        "monthly_expected": monthly_expected,
        "monthly_missed_total": monthly_missed_total,
        "iit_total": iit_total,

        # VL Metrics
        "vl_denominator": vl_denominator,
        "vl_sample_collected": vl_sample_collected,
        "vl_result_received": vl_result_received,
        "vl_coverage": vl_coverage,
        "vl_coverage_gap": vl_coverage_gap,
        "vl_suppression_rate": vl_suppression_rate,

        # EAC / AHD
        "eac_count": eac_count,
        "post_eac_vl_count": post_eac_vl_count,
        "ahd_count": ahd_count,
    }

    return render(request, "dashboard.html", context)

# ================================




@login_required
def refill_list(request):
    today = timezone.now().date()
    week_end = today + timedelta(days=7)

    # =================== GET FILTERS ===================
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    search_unique_id = request.GET.get("search_unique_id")

    facilities = Facility.objects.all()
    case_managers = (
        Refill.objects.values_list("case_manager", flat=True)
        .distinct()
        .order_by("case_manager")
    )

    refills = Refill.objects.all()

    # =================== SAFE DATE PARSING ===================
    start_date_obj = None
    end_date_obj = None
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            pass

    # =================== APPLY FILTERS ===================
    if facility_id and facility_id != "None":
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager and selected_case_manager != "None":
        refills = refills.filter(case_manager=selected_case_manager)

    if start_date_obj:
        refills = refills.filter(next_appointment__gte=start_date_obj)

    if end_date_obj:
        refills = refills.filter(next_appointment__lte=end_date_obj)

    if search_unique_id and search_unique_id.strip():
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    # =================== ADD DISPLAY FIELDS ===================
    for refill in refills:
        # Missed appointment days (temporary attribute, NOT overriding property)
        refill.days_missed_display = (
            (today - refill.next_appointment).days
            if refill.next_appointment and refill.next_appointment < today
            else 0
        )
        refill.missed_appointment = refill.days_missed_display > 0

        # Fallback dates for template
        refill.display_last_pickup = refill.last_pickup_date
        refill.display_next_appointment = refill.next_appointment

        # All other metrics now come from model properties
        # - refill.is_vl_eligible
        # - refill.is_suppressed
        # - refill.ahd
        # - refill.eac
        # - refill.eac_status
        # - refill.post_eac_vl_due
        # - refill.tpt_status

    # =================== GROUP BY PERIOD ===================
    daily_expected = refills.filter(next_appointment=today)
    weekly_expected = refills.filter(next_appointment__range=[today, week_end])
    monthly_expected = refills.filter(
        next_appointment__year=today.year,
        next_appointment__month=today.month
    )

    # =================== PAGINATION ===================
    daily_page = Paginator(daily_expected.order_by("next_appointment"), 10)
    weekly_page = Paginator(weekly_expected.order_by("next_appointment"), 10)
    monthly_page = Paginator(monthly_expected.order_by("next_appointment"), 10)

    # =================== EXCEL EXPORT ===================
    if "download" in request.GET:
        return export_refills_to_excel(refills)

    # =================== CONTEXT ===================
    context = {
        "facilities": facilities,
        "case_managers": case_managers,
        "selected_facility": facility_id,
        "selected_case_manager": selected_case_manager,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "today": today,
        "seven_days_from_today": week_end,
        "search_unique_id": search_unique_id,
        "query_params": request.GET.urlencode(),
        "periods": [
            {"name": "Daily", "page_obj": daily_page.get_page(request.GET.get("daily_page"))},
            {"name": "Weekly", "page_obj": weekly_page.get_page(request.GET.get("weekly_page"))},
            {"name": "Monthly", "page_obj": monthly_page.get_page(request.GET.get("monthly_page"))},
        ],
    }

    return render(request, "refill_list.html", context)




def export_refills_to_excel(refills):
    """
    Export Refills to Excel using the SAME dynamic logic as refill_list.
    No property is overwritten. Everything is read-only & model-safe.
    """
    today = timezone.now().date()

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
        'Remark',  # NEW Remark Column
    ]
    ws.append(headers)

    for refill in refills:

        # ---------------- NEXT APPOINTMENT & MISSED ----------------
        next_appointment = refill.next_appointment
        days_missed = (
            (today - next_appointment).days
            if next_appointment and next_appointment < today
            else 0
        )

        # ---------------- VL ELIGIBILITY ----------------
        vl_eligibility = "Eligible" if refill.is_vl_eligible else "Not Eligible"

        # ---------------- AHD ----------------
        ahd_status = "Eligible" if refill.current_art_status == "Restart" else "Not Eligible"

        # ---------------- EAC (MATCHES refill_list) ----------------
        eac_status = "Not Eligible"
        post_eac_vl_due = False

        eac_start_date = getattr(refill, "eac_start_date", None)
        eac_sessions_completed = getattr(refill, "eac_sessions_completed", 0) or 0

        if (
            refill.vl_result is not None
            and refill.vl_result >= 1000
            and refill.vl_sample_collection_date is not None
            and eac_start_date is None
        ):
            if eac_sessions_completed == 0:
                eac_status = "Eligible for 1st EAC"
            elif eac_sessions_completed == 1:
                eac_status = "Eligible for 2nd EAC"
            elif eac_sessions_completed == 2:
                eac_status = "Eligible for 3rd EAC"
            elif eac_sessions_completed >= 3:
                eac_status = "Post-EAC VL Due"
                post_eac_vl_due = True

        # ---------------- TPT STATUS (MODEL-DRIVEN) ----------------
        tpt_status = refill.tpt_status

        # ---------------- WRITE ROW ----------------
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
            eac_status,
            ahd_status,
            refill.tpt_start_date.strftime("%Y-%m-%d") if refill.tpt_start_date else "",
            refill.tpt_completion_date.strftime("%Y-%m-%d") if refill.tpt_completion_date else "",
            tpt_status,
            refill.remark or "",  # NEW Remark
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
    """
    Create a new refill or update an existing one.
    Dynamically calculates:
    - Next appointment
    - Days missed
    - VL eligibility & suppression
    - Dynamic EAC status
    - AHD status
    - TPT status
    """
    today = timezone.now().date()

    if pk:
        refill = get_object_or_404(Refill, pk=pk)
    else:
        refill = Refill()

    if request.method == "POST":
        form = RefillForm(request.POST, instance=refill)
        if form.is_valid():
            refill = form.save(commit=False)

            # ------------------ Next Appointment ------------------
            if refill.last_pickup_date and refill.months_of_refill_days:
                refill.next_appointment = refill.last_pickup_date + timedelta(days=float(refill.months_of_refill_days) * 30)
            else:
                refill.next_appointment = None

            # ------------------ Missed Appointment ------------------
            if refill.next_appointment and refill.next_appointment < today:
                refill.missed_appointment = True
                refill.days_missed = (today - refill.next_appointment).days
            else:
                refill.missed_appointment = False
                refill.days_missed = 0

            # ------------------ VL Eligibility & Suppression ------------------
            if refill.vl_result is not None:
                refill.is_vl_eligible = refill.vl_result >= 1000
                refill.is_suppressed = refill.vl_result < 1000
            else:
                refill.is_vl_eligible = False
                refill.is_suppressed = None

            # ------------------ AHD Status ------------------
            refill.ahd = (refill.current_art_status == "Restart")

            # ------------------ Dynamic EAC Status ------------------
            refill.eac = False
            refill.eac_status = "Not Eligible"
            refill.post_eac_vl_due = False

            if refill.vl_result is not None and refill.vl_result >= 1000:
                refill.eac = True
                # Determine base date for EAC (start date or last pickup)
                base_date = refill.eac_start_date or refill.last_pickup_date or today
                sessions = getattr(refill, "eac_sessions_completed", 0)

                # Calculate next session dynamically
                next_session_due = base_date + timedelta(days=(sessions + 1) * 30)
                if sessions < 3 and today >= next_session_due:
                    refill.eac_status = f"Due for EAC Session {sessions + 1}"
                elif sessions >= 3:
                    refill.post_eac_vl_due = True
                    refill.eac_status = "Post-EAC VL due"
                else:
                    refill.eac_status = "Eligible for EAC"

            # ------------------ TPT Status ------------------
            if refill.tpt_completion_date and (refill.tpt_completion_date <= today):
                refill.tpt_status = "Completed"
            elif refill.tpt_start_date:
                refill.tpt_status = "Ongoing"
            else:
                refill.tpt_status = ""

            refill.save()
            return redirect("refill_list")

    else:
        form = RefillForm(instance=refill)

    return render(
        request,
        "refill_form.html",
        {
            "form": form,
            "today": today
        }
    )







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
            return redirect('daily_refill_list')

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
    """
    refill = get_object_or_404(Refill, pk=pk)
    form = RefillForm(request.POST or None, instance=refill)

    if form.is_valid():
        refill = form.save(commit=False)

        # Auto recalculate next appointment
        if refill.last_pickup_date and refill.months_of_refill_days:
            days = float(refill.months_of_refill_days) * 30
            refill.next_appointment = refill.last_pickup_date + timedelta(days=days)

        refill.save()  # vl_status is read-only
        return redirect('refill_list')

    return render(request, "refill_form.html", {"form": form})


# ================= REFILL ADD OR UPDATE =================







@login_required
def track_refills(request):
    today = timezone.now().date()
    start_of_week = today - timedelta(days=today.weekday())
    start_of_month = today.replace(day=1)

    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    start_date_obj = end_date_obj = None
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            pass
    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            pass

    facilities = Facility.objects.all()
    refills = Refill.objects.all()

    if facility_id:
        refills = refills.filter(facility_id=facility_id)
    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)
    if start_date_obj:
        refills = refills.filter(last_pickup_date__gte=start_date_obj)
    if end_date_obj:
        refills = refills.filter(last_pickup_date__lte=end_date_obj)

    # ✅ SAFE COMPUTED VALUES
    for refill in refills:
        refill.days_missed_display = (
            (today - refill.next_appointment).days
            if refill.next_appointment and refill.next_appointment < today
            else 0
        )
        refill.missed_appointment = refill.days_missed_display > 0

    daily_qs = refills.filter(last_pickup_date=today)
    weekly_qs = refills.filter(last_pickup_date__gte=start_of_week)
    monthly_qs = refills.filter(last_pickup_date__gte=start_of_month)

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "periods": [
            ("Daily", daily_qs),
            ("Weekly", weekly_qs),
            ("Monthly", monthly_qs),
        ],
    }

    return render(request, "track_refills.html", context)



def export_track_refills_to_excel(refills):
    """
    Export refills to Excel, including VL Eligibility Status.
    """
    today = timezone.now().date()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track Refills Data"

    headers = [
        'Unique ID', 'Facility', 'Last Pickup Date', 'Refill Days', 'Sex',
        'Current Regimen', 'Case Manager', 'Next Appointment',
        'Days Missed', 'VL Eligibility Status'
    ]
    ws.append(headers)

    for refill in refills:
        refill.calculate_dates()
        next_appointment = refill.next_appointment.strftime("%Y-%m-%d") if refill.next_appointment else ""
        last_pickup = refill.last_pickup_date.strftime("%Y-%m-%d") if refill.last_pickup_date else "Never Picked"
        days_missed = (today - refill.next_appointment).days if refill.next_appointment and refill.next_appointment < today else 0
        vl_status = "Eligible" if refill.is_vl_eligible else "Not Eligible"

        row = [
            refill.unique_id,
            refill.facility.name if refill.facility else "",
            last_pickup,
            refill.months_of_refill_days,
            refill.sex,
            refill.current_regimen,
            refill.case_manager or "",
            next_appointment,
            days_missed,
            vl_status
        ]
        ws.append(row)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="Track_Refills_{today}.xlsx"'
    wb.save(response)
    return response









@login_required
def daily_refill_list(request):
    today = timezone.now().date()

    # ================= FACILITIES =================
    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()

    # ================= CASE MANAGERS =================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm.strip()})
    selected_case_manager = request.GET.get("case_manager")

    # ================= DAILY REFILLS =================
    refills = Refill.objects.filter(next_appointment=today).order_by('unique_id')

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # ================= DYNAMIC VL ELIGIBILITY =================
    for refill in refills:
        if getattr(refill, "vl_sample_collection_date", None) == today:
            refill.vl_eligible_display = False
        else:
            refill.vl_eligible_display = getattr(refill, "is_vl_eligible", False)

        # Calculate days missed temporarily (for display only)
        if refill.next_appointment and refill.next_appointment < today:
            refill.days_missed_display = (today - refill.next_appointment).days
        else:
            refill.days_missed_display = 0

    # ================= CONTEXT =================
    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "refills": refills,
    }

    return render(request, "daily_refill_list.html", context)






@login_required
def missed_refills(request):
    today = timezone.now().date()

    # ================= GET FILTER PARAMETERS =================
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
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass

    if case_manager:
        refills = refills.filter(case_manager__iexact=case_manager.strip())

    if search_unique_id:
        refills = refills.filter(unique_id__icontains=search_unique_id.strip())

    # ================= DATE FILTER =================
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
            refills = refills.filter(next_appointment__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
            refills = refills.filter(next_appointment__lte=end_date_obj)
        except ValueError:
            pass

    # ================= MISSED REFILLS LOGIC =================
    missed_list = refills.filter(
        next_appointment__lt=today
    ).order_by("next_appointment")

    # ================= CALCULATED DISPLAY FIELDS =================
    for refill in missed_list:
        if refill.next_appointment:
            days_missed = (today - refill.next_appointment).days
            refill.days_missed_display = max(days_missed, 0)

            if refill.days_missed_display >= 28:
                refill.iit_status_display = "IIT"
            elif refill.days_missed_display > 0:
                refill.iit_status_display = f"{28 - refill.days_missed_display} days to IIT"
            else:
                refill.iit_status_display = "On Track"
        else:
            refill.days_missed_display = 0
            refill.iit_status_display = "On Track"

    total_missed = missed_list.count()

    # ================= EXPORT TO EXCEL =================
    if request.GET.get("export") == "excel":
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Missed Refills"

        headers = [
            "Unique ID",
            "Case Manager",
            "Facility",
            "Last Pickup Date",
            "Next Appointment",
            "Days Missed",
            "IIT Status",
            "VL Status",
        ]
        worksheet.append(headers)

        for col in range(1, len(headers) + 1):
            worksheet.cell(row=1, column=col).font = Font(bold=True)

        for refill in missed_list:
            worksheet.append([
                refill.unique_id,
                refill.case_manager or "",
                refill.facility.name if refill.facility else "",
                refill.last_pickup_date.strftime("%Y-%m-%d") if refill.last_pickup_date else "",
                refill.next_appointment.strftime("%Y-%m-%d") if refill.next_appointment else "",
                refill.days_missed_display,
                refill.iit_status_display,
                refill.vl_status,  # MODEL PROPERTY
            ])

        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response["Content-Disposition"] = 'attachment; filename="missed_refills.xlsx"'
        workbook.save(response)
        return response

    # ================= PAGINATION =================
    paginator = Paginator(missed_list, 25)
    page_number = request.GET.get("page")
    page_obj = paginator.get_page(page_number)

    # ================= UNIQUE CASE MANAGERS =================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm and cm.strip()})

    # ================= QUERY PARAMS =================
    query_params = request.GET.copy()
    if "page" in query_params:
        query_params.pop("page")

    context = {
        "page_obj": page_obj,
        "today": today,
        "total_missed": total_missed,
        "facilities": Facility.objects.all(),
        "case_managers": case_managers,
        "selected_facility": facility_id,
        "selected_case_manager": case_manager,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "search_unique_id": search_unique_id,
        "query_params": query_params.urlencode(),
    }

    return render(request, "missed_refills.html", context)







@login_required
def track_vl(request):
    today = timezone.now().date()

    # ================== FILTERS ==================
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    selected_unique_id = request.GET.get("unique_id")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # Safe date parsing
    start_date_obj = None
    end_date_obj = None
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date_obj = None
    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            end_date_obj = None

    facilities = Facility.objects.all()
    refills = Refill.objects.all().order_by('-vl_sample_collection_date')

    # Apply filters
    if facility_id:
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass
    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)
    if selected_unique_id:
        refills = refills.filter(unique_id__icontains=selected_unique_id)
    if start_date_obj:
        refills = refills.filter(vl_sample_collection_date__gte=start_date_obj)
    if end_date_obj:
        refills = refills.filter(vl_sample_collection_date__lte=end_date_obj)

    # ================== PAGINATION ==================
    paginator = Paginator(refills, 10)  # 10 per page
    page_number = request.GET.get("page")
    vl_refills = paginator.get_page(page_number)

    # ================== CASE MANAGERS ==================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm and cm.strip()})

    # ================== EXCEL DOWNLOAD ==================
    if "download" in request.GET:
        return export_vl_to_excel(refills)

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "selected_unique_id": selected_unique_id,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "vl_refills": vl_refills,
    }

    return render(request, "track_vl.html", context)


def export_vl_to_excel(refills):
    today = timezone.now().date()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track VL"

    # Header
    headers = ["Unique ID", "Facility", "Date VL Collected", "Date Refilled", "Case Manager"]
    ws.append(headers)

    for refill in refills:
        row = [
            refill.unique_id,
            refill.facility.name if refill.facility else "",
            refill.vl_sample_collection_date.strftime("%Y-%m-%d") if refill.vl_sample_collection_date else "",
            refill.last_pickup_date.strftime("%Y-%m-%d") if refill.last_pickup_date else "",
            refill.case_manager or "",
        ]
        ws.append(row)

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="Track_VL_{today}.xlsx"'
    wb.save(response)

    return response


    

