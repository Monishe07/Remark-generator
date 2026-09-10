import pandas as pd
import re

input_file = r"C:\Users\USER\Downloads\Call Logs Report-09_09_2026 to 09_09_2026_Branch-Chennai.xls"

output_file = r"C:\Users\USER\Desktop\chennai_lead_report_with_remarks.xlsx"

# ============================================================
# 2. READ EXCEL WITHOUT ASSUMING HEADER
# ============================================================
raw = pd.read_csv(
    input_file,
    sep="\t",
    header=None
)


# ============================================================
# 3. FIND THE LOWER INBOUND / OUTBOUND HEADER
# ============================================================

header_rows = []

for i, row in raw.iterrows():

    values = [
        str(x).strip().lower()
        for x in row.tolist()
    ]

    if (
        "call type" in values
        and "call status" in values
        and "call duration & grade" in values
    ):
        header_rows.append(i)

if not header_rows:
    raise ValueError(
        "Call Type, Call status, Call Duration & Grade columns not found!"
    )

# Last matching header = lower table
header_row = header_rows[-1]

print("Lower table header found at row:", header_row)

# ============================================================
# 4. READ LOWER TABLE
# ============================================================

df = pd.read_csv(
    input_file,
    sep="\t",
    header=header_row
)
# Remove completely empty rows
df = df.dropna(how="all").copy()

# Clean column names
df.columns = [
    str(col).strip()
    for col in df.columns
]

# ============================================================
# 5. FIND REQUIRED COLUMNS
# ============================================================

def find_column(column_name):

    target = column_name.lower().strip()

    for col in df.columns:

        if str(col).lower().strip() == target:
            return col

    return None


call_type_col = find_column("Call Type")
call_status_col = find_column("Call status")
duration_col = find_column("Call Duration & Grade")


if call_type_col is None:
    raise ValueError("Call Type column not found!")

if call_status_col is None:
    raise ValueError("Call status column not found!")

if duration_col is None:
    raise ValueError(
        "Call Duration & Grade column not found!"
    )

# ============================================================
# 6. FIND CUSTOMER / PHONE NUMBER COLUMN
# ============================================================

phone_col = None

possible_phone_columns = [
    "Phone Number",
    "Phone",
    "Mobile",
    "Mobile Number",
    "Contact Number",
    "Contact",
    "Customer Number",
    "Customer Mobile"
]

for possible in possible_phone_columns:

    phone_col = find_column(possible)

    if phone_col:
        break


# If exact name is not found,
# automatically detect a column containing phone numbers.

if phone_col is None:

    for col in df.columns:

        sample = (
            df[col]
            .dropna()
            .astype(str)
            .str.replace(r"\D", "", regex=True)
        )

        if len(sample) == 0:
            continue

        valid_numbers = sample[
            sample.str.len().between(10, 12)
        ]

        if len(valid_numbers) > 0:

            phone_col = col
            break


if phone_col is None:
    raise ValueError(
        "Customer / Phone Number column could not be detected!"
    )


print("Customer number column:", phone_col)

# ============================================================
# 7. KEEP ONLY INBOUND / OUTBOUND RECORDS
# ============================================================

df["__CALL_TYPE_TEMP__"] = (
    df[call_type_col]
    .astype(str)
    .str.strip()
    .str.upper()
)

df = df[
    df["__CALL_TYPE_TEMP__"].isin(
        ["INBOUND", "OUTBOUND"]
    )
].copy()

df.drop(
    columns=["__CALL_TYPE_TEMP__"],
    inplace=True
)

# ============================================================
# 8. CLEAN CUSTOMER NUMBER
# ============================================================

def clean_phone(value):

    if pd.isna(value):
        return ""

    # Keep only numbers
    phone = re.sub(
        r"\D",
        "",
        str(value)
    )

    return phone


# ============================================================
# 9. EXTRACT ONLY CALL DURATION
# ============================================================

def extract_duration(value):

    if pd.isna(value):
        return "0 sec"

    text = str(value).strip().lower()

    # Find minutes
    minute_match = re.search(
        r"(\d+)\s*mins?",
        text
    )

    # Find seconds
    second_match = re.search(
        r"(\d+)\s*secs?",
        text
    )

    minutes = (
        int(minute_match.group(1))
        if minute_match
        else 0
    )

    seconds = (
        int(second_match.group(1))
        if second_match
        else 0
    )

    # Example:
    # 1 mins 20 sec - Bronze
    # Output:
    # 1 min 20 sec

    if minutes > 0 and seconds > 0:

        return f"{minutes} min {seconds} sec"

    elif minutes > 0:

        return f"{minutes} min"

    else:

        return f"{seconds} sec"


# ============================================================
# 10. GENERATE REMARKS
# ============================================================

# This set stores customer numbers that
# have already appeared earlier in the Excel.

seen_customers = set()

remarks = []


for index, row in df.iterrows():

    # --------------------------------------------------------
    # Get values
    # --------------------------------------------------------

    phone = clean_phone(
        row[phone_col]
    )

    call_type = str(
        row[call_type_col]
    ).strip().upper()

    status = str(
        row[call_status_col]
    ).strip().upper()

    duration = extract_duration(
        row[duration_col]
    )

    # --------------------------------------------------------
    # Check whether this customer appeared before
    # --------------------------------------------------------

    previous_call = (
        phone != ""
        and phone in seen_customers
    )

    # ========================================================
    # LOGIC 1:
    # BUSY / EXECUTIVE BUSY
    # ========================================================

    if status in [
        "BUSY",
        "EXECUTIVE BUSY"
    ]:

        # Even if customer called previously,
        # BUSY should remain NO FOLLOWUP.

        remark = "NO FOLLOWUP"

    # ========================================================
    # LOGIC 2:
    # ANSWER
    # ========================================================

    elif status == "ANSWER":

        # -----------------------------
        # INBOUND
        # -----------------------------

        if call_type == "INBOUND":

            if previous_call:

                remark = (
                    f"FOLLOWUP DONE - "
                    f"ANSWERED INBOUND ({duration})"
                )

            else:

                remark = (
                    f"ANSWERED INBOUND "
                    f"({duration})"
                )

        # -----------------------------
        # OUTBOUND
        # -----------------------------

        elif call_type == "OUTBOUND":

            if previous_call:

                remark = (
                    f"FOLLOWUP DONE - "
                    f"ANSWERED OUTBOUND ({duration})"
                )

            else:

                remark = (
                    f"ANSWERED OUTBOUND "
                    f"({duration})"
                )

        else:

            remark = f"ANSWERED ({duration})"

    # ========================================================
    # LOGIC 3:
    # NO ANSWER
    # ========================================================

    elif status in [
        "NOANSWER",
        "NO ANSWER"
    ]:

        if previous_call:

            remark = "FOLLOWUP DONE - NO ANSWER"

        else:

            remark = "NO ANSWER"

    # ========================================================
    # LOGIC 4:
    # CANCEL / CANCELLED
    # ========================================================

    elif status in [
        "CANCEL",
        "CANCELLED"
    ]:

        if previous_call:

            remark = "FOLLOWUP DONE - CANCELLED"

        else:

            remark = "CANCELLED"

    # ========================================================
    # UNKNOWN STATUS
    # ========================================================

    else:

        remark = "CHECK"

    # --------------------------------------------------------
    # Store remark
    # --------------------------------------------------------

    remarks.append(remark)

    # --------------------------------------------------------
    # IMPORTANT:
    # Add this customer to seen list AFTER processing
    # the current row.
    #
    # So first call = normal remark
    # second call = followup remark
    # --------------------------------------------------------

    if phone != "":

        seen_customers.add(phone)


# ============================================================
# 11. ADD REMARKS COLUMN
# ============================================================

df["Remarks"] = remarks


# ============================================================
# 12. SAVE OUTPUT EXCEL
# ============================================================

df.to_excel(
    output_file,
    index=False
)


# ============================================================
# 13. SUCCESS MESSAGE
# ============================================================

print("\n======================================")
print("REMARKS GENERATED SUCCESSFULLY")
print("======================================")

print("Input file :", input_file)
print("Output file:", output_file)

print("\nTotal records processed:", len(df))