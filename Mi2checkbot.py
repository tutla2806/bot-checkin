import logging
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, CallbackContext
import datetime

# Cấu hình logging
logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO)

# Token bot của bạn
TOKEN = '7839074621:AAEnNwsHhafSJGcihbgoYk7VT5grp7C8tK4'

# Đường dẫn đến file Excel
attendance_file = 'Checkin.xlsx'

# Kiểm tra nếu file Excel đã tồn tại, nếu chưa thì tạo file mới
def check_file_exists():
    try:
        df = pd.read_excel(attendance_file)
    except FileNotFoundError:
        logging.info("File Excel không tồn tại, tạo mới...")
        df = pd.DataFrame(columns=["Date", "Username"])  # Đảm bảo tạo đúng cột
        df.to_excel(attendance_file, index=False)
        logging.info("File Excel đã được tạo!")

# Hàm để điểm danh
async def check_in(update: Update, context: CallbackContext) -> None:
    user = update.message.from_user
    username = user.username  # Lấy tên người dùng
    now = datetime.datetime.now().strftime("%Y-%m-%d")  # Lấy ngày hiện tại
    logging.info(f"User {username} checked in at {now}")

    # Đọc file Excel vào DataFrame
    df = pd.read_excel(attendance_file)

    # Kiểm tra xem người dùng đã điểm danh hôm nay chưa
    if not ((df['Date'] == now) & (df['Username'] == username)).any():
        # Nếu chưa, thêm thông tin điểm danh
        new_row = pd.DataFrame({"Date": [now], "Username": [username]})
        df = pd.concat([df, new_row], ignore_index=True)  # Thêm dòng mới vào DataFrame
        df.to_excel(attendance_file, index=False)  # Lưu lại file Excel
        await update.message.reply_text(f"Bạn đã điểm danh thành công vào ngày {now}")
    else:
        await update.message.reply_text(f"{username} đã điểm danh rồi hôm nay!")

# Hàm để xuất file Excel
async def export_attendance(update: Update, context: CallbackContext) -> None:
    # Đọc file Excel vào DataFrame
    df = pd.read_excel(attendance_file)

    # Gửi file Excel cho người dùng
    with open(attendance_file, 'rb') as file:
        await update.message.reply_document(file, caption="Danh sách điểm danh")

def main():
    check_file_exists()  # Kiểm tra và tạo file Excel nếu chưa có

    application = Application.builder().token(TOKEN).build()

    # Thêm handler cho các lệnh
    application.add_handler(CommandHandler("checkin", check_in))
    application.add_handler(CommandHandler("export", export_attendance))

    # Bắt đầu bot
    logging.info("Bot is starting...")
    application.run_polling()

if __name__ == '__main__':
    main()
