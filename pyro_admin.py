import asyncio
import random
import time
from pyrogram.errors.exceptions.bad_request_400 import UsernameNotOccupied
from pyrogram.errors import UsernameInvalid, FloodWait, UsernameNotOccupied
from pyrogram.raw import functions
from slugify import slugify
import openpyxl
from pyrogram import Client, filters
from pyrogram.types import ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton

from pyro_main import main_group_send_menu, main_create

api_id = 10736822
api_hash = "3a730347f1f410c6d8491fbfaed0add9"
app_bot = Client("my_bot", api_id=api_id, api_hash=api_hash, bot_token='5784096253:AAEnhY3_m2WLalGWStmkYD5R2DQixbFsmyo')
app_user = Client("my_account", api_id=api_id, api_hash=api_hash)
file_ = openpyxl.load_workbook('list.xlsx')
sheet_obj_ = file_.active

data_ = [({
    'channel': sheet_obj_.cell(row=i + 1, column=1).value,
    'category': sheet_obj_.cell(row=i + 1, column=2).value,
    'name_new': sheet_obj_.cell(row=i + 1, column=4).value,
    'id': sheet_obj_.cell(row=i + 1, column=5).value,
}) for i in range(sheet_obj_.max_row) if sheet_obj_.cell(row=i + 1, column=5).value]
channel_ids = [channel_id.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_') for
               channel_id in data_]
file_.close()


async def send_new_msg(app_bot,app_user,message,new_channel_id):
    last_mes = app_user.get_messages(message.chat.id,message.id-1)
    mg_id = app_user.get_messages(message.chat.id,message.id-1).media_group_id
    if (message.photo or last_mes.photo or last_mes.media_group_id) and mg_id != message.media_group_id:
        kk = random.randrange(1, 4)
        await asyncio.sleep(kk)
        if message.media_group_id:
            try:
                await app_user.copy_media_group(new_channel_id,
                                                message.chat.id, message.id)
            except FloodWait as e:
                print('wait for', e.value, 'to send message')
                await asyncio.sleep(e.value + 2)
                await app_user.copy_media_group(new_channel_id,
                                                message.chat.id, message.id)
            except Exception as e:
                print(e)
        elif message.photo:
            try:
                await app_user.send_photo(new_channel_id, message.photo.file_id)
            except FloodWait as e:
                print('wait for', e.value, 'to send message')
                await asyncio.sleep(e.value + 2)
                await app_user.send_photo(new_channel_id, message.photo.file_id)
            except Exception as e:
                print(e)
        if message.text:
            try:
                await app_user.send_message(new_channel_id, message.text)
                print('text')
            except FloodWait as e:
                print('wait for', e.value, 'to send message')
                await asyncio.sleep(e.value + 2)
                await app_user.send_message(new_channel_id, message.text)
            except Exception as e:
                print(e)


@app_bot.on_message(filters.command('start') & filters.private)
async def on_start(app_bot, message):
    await app_bot.send_message(message.chat.id, 'started bot')


@app_bot.on_message(filters.command('send_menu') & filters.private)
async def on_start(app_bot, message):
    await main_group_send_menu(app_bot, app_user)


@app_bot.on_message(filters.command('create_channels') & filters.private)
async def on_start(app_bot, message):
    await main_create(app_bot, app_user)


@app_bot.on_message(filters.channel and filters.create(lambda self, c, m: (m.chat.username in channel_ids)))
async def check_updates(app_bot, message):
    new_channel_id = 0
    for channel in data_:
        if channel.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_') == message.chat.id:
            new_channel_id = channel.get('id')
    await send_new_msg(app_bot,app_user,message,new_channel_id)


app_user.start()
app_bot.run()
