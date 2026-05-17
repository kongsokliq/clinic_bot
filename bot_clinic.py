import os
from openpyxl import Workbook, load_workbook

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup
)

from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters
)

# =====================================================
# CONFIG
# =====================================================

TOKEN = "8703254616:AAE4PRrjID7duF4V0g01CY7OOSkObKiO3x0"

ADMIN_IDS = [387739135]

FILE_NAME = "patients.xlsx"

IMAGE_FOLDER = "patient_images"

# =====================================================
# CREATE IMAGE FOLDER
# =====================================================

os.makedirs(IMAGE_FOLDER, exist_ok=True)

# =====================================================
# CREATE EXCEL
# =====================================================

def init_excel():

    if not os.path.exists(FILE_NAME):

        wb = Workbook()
        ws = wb.active

        ws.title = "Patients"

        headers = [
            "Name",
            "Age",
            "Gender",
            "Phone",
            "Address",
            "Diagnosis",
            "Treatment",
            "VisitDate",
            "Note",
            "Image"
        ]

        ws.append(headers)

        wb.save(FILE_NAME)

# =====================================================
# CHECK ADMIN
# =====================================================

def is_admin(user_id):

    return user_id in ADMIN_IDS

# =====================================================
# SAVE PATIENT
# =====================================================

def save_patient(data):

    wb = load_workbook(FILE_NAME)
    ws = wb["Patients"]

    ws.append(data)

    wb.save(FILE_NAME)

# =====================================================
# GET ALL PATIENTS
# =====================================================

def get_all_patients():

    wb = load_workbook(FILE_NAME)
    ws = wb["Patients"]

    patients = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        patients.append(row)

    return patients

# =====================================================
# SEARCH PATIENT
# =====================================================

def search_patient(keyword):

    wb = load_workbook(FILE_NAME)
    ws = wb["Patients"]

    results = []

    for row in ws.iter_rows(min_row=2, values_only=True):

        if keyword.lower() in str(row[0]).lower():
            results.append(row)

    return results

# =====================================================
# DELETE PATIENT
# =====================================================

def delete_patient(name):

    wb = load_workbook(FILE_NAME)
    ws = wb["Patients"]

    deleted = False

    for row in range(2, ws.max_row + 1):

        patient_name = ws.cell(row=row, column=1).value

        if str(patient_name).lower() == name.lower():

            image_path = ws.cell(row=row, column=10).value

            if image_path and os.path.exists(image_path):
                os.remove(image_path)

            ws.delete_rows(row)

            deleted = True
            break

    wb.save(FILE_NAME)

    return deleted

# =====================================================
# SAVE IMAGE PATH
# =====================================================

def save_image_to_excel(name, image_path):

    wb = load_workbook(FILE_NAME)
    ws = wb["Patients"]

    for row in range(2, ws.max_row + 1):

        patient_name = ws.cell(row=row, column=1).value

        if str(patient_name).lower() == name.lower():

            ws.cell(row=row, column=10).value = image_path

            wb.save(FILE_NAME)

            return True

    return False

# =====================================================
# MAIN MENU
# =====================================================

def main_menu():

    keyboard = [

        [InlineKeyboardButton(
            "➕ Add Patient",
            callback_data="add_patient"
        )],

        [InlineKeyboardButton(
            "📋 Patient List",
            callback_data="list_patients"
        )],

        [InlineKeyboardButton(
            "🔍 Search Patient",
            callback_data="search_patient"
        )],

        [InlineKeyboardButton(
            "📷 Upload Image",
            callback_data="upload_image"
        )],

        [InlineKeyboardButton(
            "🗑 Delete Patient",
            callback_data="delete_patient"
        )],

        [InlineKeyboardButton(
            "📥 Download Excel",
            callback_data="download_excel"
        )]
    ]

    return InlineKeyboardMarkup(keyboard)

# =====================================================
# START
# =====================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user_id = update.effective_user.id

    if not is_admin(user_id):

        await update.message.reply_text(
            "❌ Access denied"
        )
        return

    await update.message.reply_text(
        "🏥 Clinic Management System",
        reply_markup=main_menu()
    )

# =====================================================
# BUTTON ROUTER
# =====================================================

async def button_router(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    data = query.data

    # ADD PATIENT
    if data == "add_patient":

        context.user_data["action"] = "save_patient"

        await query.message.reply_text(
            "Send:\n\n"
            "name,age,gender,phone,address,diagnosis,treatment,visitdate,note"
        )

    # LIST PATIENTS
    elif data == "list_patients":

        patients = get_all_patients()

        if not patients:

            text = "❌ No patients"

        else:

            text = "📋 Patient List\n\n"

            for p in patients:

                text += (
                    f"👤 {p[0]}\n"
                    f"📞 {p[3]}\n"
                    f"🩺 {p[5]}\n"
                    f"📅 {p[7]}\n\n"
                )

        await query.message.reply_text(text)

    # SEARCH
    elif data == "search_patient":

        context.user_data["action"] = "search_patient"

        await query.message.reply_text(
            "🔍 Send patient name"
        )

    # DELETE
    elif data == "delete_patient":

        context.user_data["action"] = "delete_patient"

        await query.message.reply_text(
            "🗑 Send patient name"
        )

    # UPLOAD IMAGE
    elif data == "upload_image":

        context.user_data["action"] = "upload_image_name"

        await query.message.reply_text(
            "📷 Send patient name first"
        )

    # DOWNLOAD EXCEL
    elif data == "download_excel":

        await query.message.reply_document(
            document=open(FILE_NAME, "rb"),
            filename="patients.xlsx"
        )

# =====================================================
# HANDLE TEXT
# =====================================================

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user_id = update.effective_user.id

    if not is_admin(user_id):
        return

    action = context.user_data.get("action")

    text = update.message.text.strip()

    # =================================================
    # SAVE PATIENT
    # =================================================

    if action == "save_patient":

        try:

            data = [x.strip() for x in text.split(",")]

            if len(data) != 9:
                raise Exception()

            data.append("")

            save_patient(data)

            await update.message.reply_text(
                "✅ Patient saved"
            )

        except:

            await update.message.reply_text(
                "❌ Invalid format"
            )

        context.user_data["action"] = None

    # =================================================
    # SEARCH
    # =================================================

    elif action == "search_patient":

        results = search_patient(text)

        if not results:

            await update.message.reply_text(
                "❌ Not found"
            )

        else:

            msg = ""

            for r in results:

                msg += (
                    f"👤 Name: {r[0]}\n"
                    f"🎂 Age: {r[1]}\n"
                    f"⚧ Gender: {r[2]}\n"
                    f"📞 Phone: {r[3]}\n"
                    f"🏠 Address: {r[4]}\n"
                    f"🩺 Diagnosis: {r[5]}\n"
                    f"💊 Treatment: {r[6]}\n"
                    f"📅 Visit: {r[7]}\n"
                    f"📝 Note: {r[8]}\n\n"
                )

            await update.message.reply_text(msg)

        context.user_data["action"] = None

    # =================================================
    # DELETE
    # =================================================

    elif action == "delete_patient":

        success = delete_patient(text)

        if success:

            await update.message.reply_text(
                "✅ Deleted successfully"
            )

        else:

            await update.message.reply_text(
                "❌ Patient not found"
            )

        context.user_data["action"] = None

    # =================================================
    # IMAGE NAME
    # =================================================

    elif action == "upload_image_name":

        context.user_data["patient_name"] = text

        context.user_data["action"] = "waiting_photo"

        await update.message.reply_text(
            "📤 Send image now"
        )

# =====================================================
# HANDLE PHOTO
# =====================================================

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):

    action = context.user_data.get("action")

    if action != "waiting_photo":
        return

    patient_name = context.user_data.get("patient_name")

    photo = update.message.photo[-1]

    file = await photo.get_file()

    filename = f"{patient_name}.jpg"

    filepath = os.path.join(
        IMAGE_FOLDER,
        filename
    )

    await file.download_to_drive(filepath)

    success = save_image_to_excel(
        patient_name,
        filepath
    )

    if success:

        await update.message.reply_text(
            "✅ Image uploaded"
        )

    else:

        await update.message.reply_text(
            "❌ Patient not found"
        )

    context.user_data["action"] = None

# =====================================================
# MAIN
# =====================================================

def main():

    init_excel()

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(
        CommandHandler("start", start)
    )

    app.add_handler(
        CallbackQueryHandler(button_router)
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

    print("🤖 Bot Running 24/24...")

    app.run_polling()

# =====================================================
# START BOT
# =====================================================

if __name__ == "__main__":
    main()