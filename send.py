import logging
import shutil
import os
from openpyxl import load_workbook
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from atanish import take_attendance, extrat_stu

API_TOKEN = '6393650795:AAG8ISi3VT-RN-EnfoJ5G2JuVegoWhg0hNc'  # Replace with your actual API token

logging.basicConfig(level=logging.INFO)

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)
dp.middleware.setup(LoggingMiddleware())

students = extrat_stu("attendance.xlsx")  # Predefined list of students
student_selections = {}

@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    global students, student_selections

    student_selections = {student: None for student in students}

    keyboard = types.InlineKeyboardMarkup()
    keyboard.add(types.InlineKeyboardButton(text="Present", callback_data="present"))
    keyboard.add(types.InlineKeyboardButton(text="Absent", callback_data="absent"))

    keyboard1 = types.InlineKeyboardMarkup()
    keyboard1.add(types.InlineKeyboardButton(text="Submit", callback_data="submit"))

    for student in students:
        await message.answer(f"Is {student} present or absent?", reply_markup=keyboard)
    await message.answer(f"Submit here", reply_markup=keyboard1)

@dp.callback_query_handler(lambda c: c.data in ["present", "absent"])
async def process_callback(callback_query: types.CallbackQuery):
    global students, student_selections

    student_name = next(student for student in students if student in callback_query.message.text)

    if callback_query.data == "present":
        student_selections[student_name] = "Present"
    else:
        student_selections[student_name] = "Absent"

    keyboard = types.InlineKeyboardMarkup()
    keyboard.add(types.InlineKeyboardButton(text="Edit", callback_data="edit"))

    new_message_text = callback_query.message.text + f"\nSelection: {student_selections[student_name]}"

    await bot.edit_message_text(chat_id=callback_query.message.chat.id,
                                message_id=callback_query.message.message_id,
                                text=new_message_text,
                                reply_markup=keyboard)

@dp.callback_query_handler(lambda c: c.data == "edit")
async def edit_callback(callback_query: types.CallbackQuery):
    global students, student_selections

    student_name = next(student for student in students if student in callback_query.message.text)

    keyboard = types.InlineKeyboardMarkup()
    keyboard.add(types.InlineKeyboardButton(text="Present", callback_data="present"))
    keyboard.add(types.InlineKeyboardButton(text="Absent", callback_data="absent"))

    await bot.edit_message_text(chat_id=callback_query.message.chat.id,
                                message_id=callback_query.message.message_id,
                                text=f"Is {student_name} present or absent?",
                                reply_markup=keyboard)

@dp.callback_query_handler(lambda c: c.data == "submit")
async def submit_callback(callback_query: types.CallbackQuery):
    global students, student_selections

    present_students = [student for student, selection in student_selections.items() if selection == "Present"]
    absent_students = [student for student, selection in student_selections.items() if selection == "Absent"]

    present_list = "\n".join(present_students)
    absent_list = "\n".join(absent_students)

    result_message = f"Present Students:\n{present_list}\n\nAbsent Students:\n{absent_list}"
    file_path = "attendance.xlsx"
    take_attendance(file_path,present_students,absent_students)

    # Create a copy of the Excel file
    original_file_path = "attendance.xlsx"
    copy_file_path = "attendance_copy.xlsx"
    shutil.copy(original_file_path, copy_file_path)

    # Modify the copy (if needed)
    # For example, you can update the copy with attendance information

    # Send the copy of the Excel file to the user
    with open(copy_file_path, 'rb') as file:
        await bot.send_document(callback_query.from_user.id, file, caption="Attendance Report")

    # Delete the copy file (optional)
    os.remove(copy_file_path)

    await bot.send_message(callback_query.from_user.id, result_message)

if __name__ == '__main__':
    from aiogram import executor
    executor.start_polling(dp, skip_updates=True)
