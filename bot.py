from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
import asyncio
from parser import parsing, url

API_TOKEN = "7911812987:AAHpm_K4N2DKRBKkDaNOCMkXCcY1MIINwgc"

bot = Bot(token=API_TOKEN)
dp = Dispatcher()

@dp.message(Command("start"))
async def send_welcome(message: types.Message):
    await message.answer("Пошёл нахуй, я не доделан")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())