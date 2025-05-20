import pandas as pd
from datetime import datetime
import streamlit as st

# --- Helper functions ---
def normalize_time(raw):
    if pd.isna(raw):
        return None
    raw = raw.lower().strip()
    ampm = 'AM' if 'a' in raw else 'PM'
    digits = raw.replace('a', '').replace('p', '')
    if len(digits) == 3:
        hour, minute = int(digits[0]), int(digits[1:])
    else:
        hour, minute = int(digits[:2]), int(digits[2:])
    return datetime.strptime(f'{hour}:{minute} {ampm}', '%I:%M %p').time()

def expected_start(clock_in):
    if clock_in <= datetime.strptime('08:00 AM', '%I:%M %p').time():
        return datetime.strptime('07:15 AM', '%I:%M %p').time()
    elif clock_in <= datetime.strptime('01:00 PM', '%I:%M %p').time():
        return datetime.strptime('12:00 PM', '%I:%M %p').time()
    else:
        return datetime.strptime('02:00 PM', '%I:%M %p').time()

def minutes_late(actual, expected):
    if not actual or not expected:
        return None
    a = datetime.strptime(actual.strftime('%H:%M'), '%H:%M')
    e = datetime.strptime(expected.strftime('%H:%M'), '%H:%M')
    return max(int((a - e).total_seconds() / 60), 0)

# --- Flexible weekday date matcher ---
def is_date_header(cell):
    if not isinstance(cell, str):
        return False
    clean = cell.strip().lower()
    return any(clean.startswith(d) for d in ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'])

# --- Main timecard processor ---
def process_timecard(file):
    df = pd.read_excel(file, sheet_name=0)
    result = []

    for i in range(len(df) - 1):
        row = df.iloc[i]
        next_row = df.iloc[i + 1]

        if any(is_date_header(cell) for cell in row):
            print(f"✅ Found date row at index {i}: {list(row)}")

            for col_index, cell in enumerate(row):
                if is_date_header(cell):
                    try:
                        day_str, date_str = cell.strip().split(' ')
                        full_date = datetime.strptime(f'{date_str}/2024', '%m/%d/%Y').strftime('%m/%d/%Y')

                        entry = next_row.iloc[col_index]
                        status = ''
                        clock_in_display = ''
                        late_by = ''

                        if pd.isna(entry):
                            status = "Missing"
                        else:
                            entry_str = str(entry).strip().lower()

                            if 'pto' in entry_str:
                                status = 'PTO'
                                clock_in_display = 'PTO'
                            elif 'holiday' in entry_str:
                                status = 'Holiday'
                                clock_in_display = 'Holiday'
                            elif '-' in entry_str and any(x in entry_str for x in ['a', 'p']):
                                time_part = entry_str.split('-')[0]
                                clock_in = normalize_time(time_part)
                                expected = expected_start(clock_in)
                                late = minutes_late(clock_in, expected)
                                status = 'Late' if late > 0 else 'On Time'
                                clock_in_display = clock_in.strftime('%H:%M')
                                late_by = late
                            else:
                                status = "Other"
                                clock_in_display = entry_str

                        result.append({
                            'Date': full_date,
                            'Clock In Time': clock_in_display,
                            'Minutes Late': late_by,
                            'Status': status
                        })

                    except Exception as e:
                        print(f"❌ Error parsing column {col_index} in row {i}: {e}")
                        continue

    df_result = pd.DataFrame(result)

    if not df_result.empty and 'Status' in df_result.columns:
        worked_shifts = df_result[df_result['Status'].isin(['Late', 'On Time'])].shape[0]
        late_shifts = df_result[df_result['Status'] == 'Late'].shape[0]
        percent_late = (late_shifts / worked_shifts) * 100 if worked_shifts > 0 else 0

        summary_row_1 = {'Date': '', 'Clock In Time': '', 'Minutes Late': '', 'Status': f'Shifts Worked: {worked_shifts}'}
        summary_row_2 = {'Date': '', 'Clock In Time': '', 'Minutes Late': '', 'Status': f'% Late: {percent_late:.1f}%'}

        df_result = pd.concat([df_result, pd.DataFrame([summary_row_1, summary_row_2])], ignore_index=True)

    return df_result

# --- Streamlit App Interface ---
st.title("⏰ Timecard Checker")

uploaded = st.file_uploader("Upload your Excel timecard (.xlsx)", type=["xlsx"])

if uploaded:
    report = process_timecard(uploaded)
    st.dataframe(report)

    if not report.empty:
        csv = report.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download CSV Report",
            data=csv,
            file_name="timecard_report.csv",
            mime="text/csv"
        )
