# =========================================================
# PROFESSIONAL CLINIC MANAGEMENT BOT
# RENDER WEBHOOK VERSION (STABLE)
# =========================================================

# INSTALL:
# pip install python-telegram-bot==20.7 openpyxl

# RENDER START COMMAND:
# python bot.py

# =========================================================
# IMPORT
# =========================================================

import os

from datetime import datetime

from openpyxl import Workbook, load_workbook

from telegram import (
    Update,
    ReplyKeyboardMarkup
)

from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters
)

# =========================================================
# TOKEN & WEBHOOK
# =========================================================

TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN:
    TOKEN = "8703254616:AAEsHBheplvD4yNT5hZSuvj5ZkRYGpkzzLQ"

RENDER_URL = os.getenv(
    "RENDER_EXTERNAL_URL"
)

PORT = int(
    os.environ.get("PORT", 10000)
)

if not TOKEN:

    raise ValueError(
        "❌ BOT_TOKEN NOT FOUND"
    )

if not RENDER_URL:

    raise ValueError(
        "❌ RENDER_EXTERNAL_URL NOT FOUND"
    )

# =========================================================
# ADMIN
# =========================================================

ADMIN_IDS = [
    387739135
]

# =========================================================
# FILES
# =========================================================

FILE_NAME = "clinic_data.xlsx"

IMAGE_FOLDER = "patient_images"

os.makedirs(
    IMAGE_FOLDER,
    exist_ok=True
)

# =========================================================
# CREATE EXCEL
# =========================================================

def init_excel():

    if not os.path.exists(FILE_NAME):

        wb = Workbook()

        ws = wb.active

        ws.title = "Patients"

        ws.append([
            "No",
            "Name",
            "Age",
            "Gender",
            "Phone",
            "Address",
            "Diagnosis",
            "Treatment",
            "MedicineList",
            "VisitDate",
            "AppointmentDate",
            "FollowUp",
            "Note",
            "Image"
        ])

        wb.save(FILE_NAME)

# =========================================================
# CHECK ADMIN
# =========================================================

def is_admin(user_id):

    return user_id in ADMIN_IDS

# =========================================================
# OPEN EXCEL
# =========================================================

def get_sheet():

    wb = load_workbook(FILE_NAME)

    ws = wb["Patients"]

    return wb, ws

# =========================================================
# DATE VALIDATION
# =========================================================

def validate_date(date_text):

    try:

        datetime.strptime(
            date_text,
            "%d.%m.%Y"
        )

        return True

    except:

        return False

# =========================================================
# NEXT NUMBER
# =========================================================

def get_next_no():

    wb, ws = get_sheet()

    return ws.max_row

# =========================================================
# SAVE PATIENT
# =========================================================

def save_patient(data):

    wb, ws = get_sheet()

    ws.append(data)

    wb.save(FILE_NAME)

# =========================================================
# GET PATIENTS
# =========================================================

def get_all_patients():

    wb, ws = get_sheet()

    patients = []

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        patients.append(row)

    return patients

# =========================================================
# SEARCH PATIENT
# =========================================================

def search_patient(keyword):

    wb, ws = get_sheet()

    keyword = keyword.lower()

    results = []

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        name = str(row[1]).lower()

        phone = str(row[4]).lower()

        if keyword in name or keyword in phone:

            results.append(row)

    return results

# =========================================================
# DELETE PATIENT
# =========================================================

def delete_patient(keyword):

    wb, ws = get_sheet()

    keyword = keyword.lower()

    deleted = False

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = str(
            ws.cell(
                row=row,
                column=2
            ).value
        ).lower()

        phone = str(
            ws.cell(
                row=row,
                column=5
            ).value
        ).lower()

        if keyword in name or keyword in phone:

            image_path = ws.cell(
                row=row,
                column=14
            ).value

            if image_path:

                if os.path.exists(image_path):

                    os.remove(image_path)

            ws.delete_rows(row)

            deleted = True

            break

    wb.save(FILE_NAME)

    return deleted

# =========================================================
# FOLLOW UP
# =========================================================

def update_followup(
    keyword,
    followup,
    medicine,
    visit_date,
    note
):

    wb, ws = get_sheet()

    keyword = keyword.lower()

    updated = False

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = str(
            ws.cell(
                row=row,
                column=2
            ).value
        ).lower()

        phone = str(
            ws.cell(
                row=row,
                column=5
            ).value
        ).lower()

        if keyword in name or keyword in phone:

            ws.cell(
                row=row,
                column=12
            ).value = followup

            ws.cell(
                row=row,
                column=9
            ).value = medicine

            ws.cell(
                row=row,
                column=10
            ).value = visit_date

            ws.cell(
                row=row,
                column=13
            ).value = note

            updated = True

            break

    wb.save(FILE_NAME)

    return updated

# =========================================================
# SET APPOINTMENT
# =========================================================

def set_appointment(
    keyword,
    appointment_date
):

    wb, ws = get_sheet()

    keyword = keyword.lower()

    updated = False

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = str(
            ws.cell(
                row=row,
                column=2
            ).value
        ).lower()

        phone = str(
            ws.cell(
                row=row,
                column=5
            ).value
        ).lower()

        if keyword in name or keyword in phone:

            ws.cell(
                row=row,
                column=11
            ).value = appointment_date

            updated = True

            break

    wb.save(FILE_NAME)

    return updated

# =========================================================
# TODAY APPOINTMENTS
# =========================================================

def today_appointments():

    wb, ws = get_sheet()

    today = datetime.now().strftime(
        "%d.%m.%Y"
    )

    results = []

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if str(row[10]) == today:

            results.append(row)

    return results

# =========================================================
# SAVE IMAGE
# =========================================================

def save_image_path(
    keyword,
    image_path
):

    wb, ws = get_sheet()

    keyword = keyword.lower()

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = str(
            ws.cell(
                row=row,
                column=2
            ).value
        ).lower()

        phone = str(
            ws.cell(
                row=row,
                column=5
            ).value
        ).lower()

        if keyword in name or keyword in phone:

            ws.cell(
                row=row,
                column=14
            ).value = image_path

            wb.save(FILE_NAME)

            return True

    return False

# =========================================================
# MENU
# =========================================================

def reply_menu():

    keyboard = [

        [
            "➕ Add Patient",
            "📋 Patient List"
        ],

        [
            "🔍 Search Patient",
            "📝 Follow Up"
        ],

        [
            "📅 Appointments",
            "📅 Today Appointments"
        ],

        [
            "📷 Upload Image",
            "🗑 Delete Patient"
        ],

        [
            "📥 Download Excel"
        ]
    ]

    return ReplyKeyboardMarkup(
        keyboard,
        resize_keyboard=True
    )

# =========================================================
# START
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if not is_admin(
        update.effective_user.id
    ):

        await update.message.reply_text(
            "❌ Access Denied"
        )

        return

    await update.message.reply_text(
        "🏥 Clinic Management Bot",
        reply_markup=reply_menu()
    )

# =========================================================
# HANDLE MESSAGE
# =========================================================

async def handle_message(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if not is_admin(
        update.effective_user.id
    ):

        return

    text = update.message.text.strip()

    action = context.user_data.get(
        "action"
    )

    # =====================================================
    # MENU
    # =====================================================

    if text == "➕ Add Patient":

        context.user_data["action"] = "add_patient"

        await update.message.reply_text(
            "Send:\n\n"
            "name,age,gender,phone,address,"
            "diagnosis,treatment,"
            "medicine_list,visit_date\n\n"
            "Date Format:\n"
            "18.05.2026"
        )

        return

    elif text == "📋 Patient List":

        patients = get_all_patients()

        if not patients:

            await update.message.reply_text(
                "❌ No Patients"
            )

            return

        msg = "📋 Patient List\n\n"

        for p in patients:

            msg += (
                f"🔢 No: {p[0]}\n"
                f"👤 {p[1]}\n"
                f"📞 {p[4]}\n"
                f"📅 Visit: {p[9]}\n\n"
            )

        await update.message.reply_text(
            msg
        )

        return

    elif text == "🔍 Search Patient":

        context.user_data["action"] = "search_patient"

        await update.message.reply_text(
            "Send patient name or phone"
        )

        return

    elif text == "📝 Follow Up":

        context.user_data["action"] = "follow_up"

        await update.message.reply_text(
            "Send:\n\n"
            "name_or_phone,followup,"
            "medicine_list,visit_date,note"
        )

        return

    elif text == "📅 Appointments":

        context.user_data["action"] = "appointments"

        await update.message.reply_text(
            "Send:\n\n"
            "name_or_phone,appointment_date\n\n"
            "Example:\n"
            "012345678,18.05.2026"
        )

        return

    elif text == "📅 Today Appointments":

        appointments = today_appointments()

        if not appointments:

            await update.message.reply_text(
                "❌ No Appointments Today"
            )

            return

        msg = "📅 Today Appointments\n\n"

        for a in appointments:

            msg += (
                f"👤 {a[1]}\n"
                f"📞 {a[4]}\n"
                f"📅 {a[10]}\n\n"
            )

        await update.message.reply_text(
            msg
        )

        return

    elif text == "🗑 Delete Patient":

        context.user_data["action"] = "delete_patient"

        await update.message.reply_text(
            "Send patient name or phone"
        )

        return

    elif text == "📷 Upload Image":

        context.user_data["action"] = "upload_image"

        await update.message.reply_text(
            "Send patient name or phone"
        )

        return

    elif text == "📥 Download Excel":

        await update.message.reply_document(
            document=open(FILE_NAME, "rb"),
            filename="clinic_data.xlsx"
        )

        return

    # =====================================================
    # ADD PATIENT
    # =====================================================

    if action == "add_patient":

        try:

            data = [
                x.strip()
                for x in text.split(",")
            ]

            if len(data) != 9:

                raise ValueError()

            if not validate_date(data[8]):

                await update.message.reply_text(
                    "❌ Invalid Date"
                )

                return

            no = get_next_no()

            full_data = [
                no
            ] + data + [
                "",
                "",
                "",
                ""
            ]

            save_patient(full_data)

            await update.message.reply_text(
                f"✅ Patient Saved\n\n"
                f"🔢 No: {no}"
            )

        except:

            await update.message.reply_text(
                "❌ Invalid Format"
            )

        context.user_data["action"] = None

    # =====================================================
    # SEARCH
    # =====================================================

    elif action == "search_patient":

        results = search_patient(text)

        if not results:

            await update.message.reply_text(
                "❌ Patient Not Found"
            )

        else:

            for r in results:

                msg = (
                    f"🔢 No: {r[0]}\n"
                    f"👤 {r[1]}\n"
                    f"📞 {r[4]}\n"
                    f"🩺 {r[6]}\n"
                    f"💉 {r[7]}\n"
                    f"💊 {r[8]}\n"
                    f"📅 Visit: {r[9]}\n"
                    f"📆 Appointment: {r[10]}\n"
                    f"📝 Follow Up: {r[11]}\n"
                    f"📌 Note: {r[12]}"
                )

                await update.message.reply_text(
                    msg
                )

        context.user_data["action"] = None

    # =====================================================
    # FOLLOW UP
    # =====================================================

    elif action == "follow_up":

        try:

            data = [
                x.strip()
                for x in text.split(",")
            ]

            if len(data) != 5:

                raise ValueError()

            success = update_followup(
                data[0],
                data[1],
                data[2],
                data[3],
                data[4]
            )

            if success:

                await update.message.reply_text(
                    "✅ Follow Up Updated"
                )

            else:

                await update.message.reply_text(
                    "❌ Patient Not Found"
                )

        except:

            await update.message.reply_text(
                "❌ Invalid Format"
            )

        context.user_data["action"] = None

    # =====================================================
    # APPOINTMENT
    # =====================================================

    elif action == "appointments":

        try:

            data = [
                x.strip()
                for x in text.split(",")
            ]

            if len(data) != 2:

                raise ValueError()

            if not validate_date(data[1]):

                await update.message.reply_text(
                    "❌ Invalid Date"
                )

                return

            success = set_appointment(
                data[0],
                data[1]
            )

            if success:

                await update.message.reply_text(
                    "✅ Appointment Saved"
                )

            else:

                await update.message.reply_text(
                    "❌ Patient Not Found"
                )

        except:

            await update.message.reply_text(
                "❌ Invalid Format"
            )

        context.user_data["action"] = None

    # =====================================================
    # DELETE
    # =====================================================

    elif action == "delete_patient":

        success = delete_patient(text)

        if success:

            await update.message.reply_text(
                "✅ Patient Deleted"
            )

        else:

            await update.message.reply_text(
                "❌ Patient Not Found"
            )

        context.user_data["action"] = None

    # =====================================================
    # UPLOAD IMAGE
    # =====================================================

    elif action == "upload_image":

        context.user_data[
            "patient_keyword"
        ] = text

        context.user_data[
            "action"
        ] = "waiting_photo"

        await update.message.reply_text(
            "📤 Send image now"
        )

# =========================================================
# HANDLE PHOTO
# =========================================================

async def handle_photo(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    action = context.user_data.get(
        "action"
    )

    if action != "waiting_photo":

        return

    keyword = context.user_data.get(
        "patient_keyword"
    )

    photo = update.message.photo[-1]

    file = await photo.get_file()

    filename = f"{keyword}.jpg"

    filepath = os.path.join(
        IMAGE_FOLDER,
        filename
    )

    await file.download_to_drive(
        filepath
    )

    success = save_image_path(
        keyword,
        filepath
    )

    if success:

        await update.message.reply_text(
            "✅ Image Uploaded"
        )

    else:

        await update.message.reply_text(
            "❌ Patient Not Found"
        )

    context.user_data["action"] = None

# =========================================================
# MAIN
# =========================================================

def main():

    init_excel()

    print("🤖 Bot Starting Webhook...")

    app = ApplicationBuilder().token(
        TOKEN
    ).build()

    app.add_handler(
        CommandHandler(
            "start",
            start
        )
    )

    app.add_handler(
        MessageHandler(
            filters.PHOTO,
            handle_photo
        )
    )

    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            handle_message
        )
    )

    print("✅ Webhook Running...")

    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        webhook_url=f"{RENDER_URL}/{TOKEN}",
        drop_pending_updates=True
    )

# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    main()