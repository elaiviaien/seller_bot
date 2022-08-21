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

from pyro_main import main_group_send_menu, main_create, send_new_msg, find_first_empty, add_new_category

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
}) for i in range(1, find_first_empty(sheet_obj_)) if sheet_obj_.cell(row=i + 1, column=5).value]
channel_ids = [channel_id.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_') for
               channel_id in data_]

file_.close()


def return_channels_ids():
    file_c = openpyxl.load_workbook('list.xlsx')
    sheet_obj_c = file_c.active

    data_c = [({
        'channel': sheet_obj_c.cell(row=i + 1, column=1).value,
        'category': sheet_obj_c.cell(row=i + 1, column=2).value,
        'name_new': sheet_obj_c.cell(row=i + 1, column=4).value,
        'id': sheet_obj_c.cell(row=i + 1, column=5).value,
    }) for i in range(1, find_first_empty(sheet_obj_c)-1) if sheet_obj_.cell(row=i + 1, column=5).value]
    channel_ids = [channel_id.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_') for
                   channel_id in data_c]
    file_c.close()
    return channel_ids


@app_bot.on_message(filters.command('start') & filters.private)
async def on_start(app_bot, message):
    await app_bot.send_message(message.chat.id, 'started bot')


@app_bot.on_message(filters.command('send_menu') & filters.private)
async def on_start(app_bot, message):
    await main_group_send_menu(app_bot, app_user)


@app_bot.on_message(filters.command('create_channels') & filters.private)
async def on_start(app_bot, message):
    await main_create(app_bot, app_user)


@app_bot.on_message(filters.command('join') & filters.private)
async def join_all(app_bot, message):
    for channel in return_channels_ids():
        await app_user.join_chat(
            channel)


@app_user.on_message(filters.channel and filters.create(lambda self, c, m: (m.chat.username in return_channels_ids())))
async def check_updates(app_bot, message):
    new_channel_id = 0
    for channel in data_:
        if channel.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-',
                                                                                                '_') == message.chat.id:
            new_channel_id = channel.get('id')
    await send_new_msg(app_bot, app_user, message, new_channel_id)


@app_bot.on_message(filters.command('add_category') & filters.private)
async def add_category(app_bot, message):
    await add_new_category(app_bot, message)

app_user.start()
app_bot.run()
