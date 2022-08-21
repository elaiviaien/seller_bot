import asyncio
import random
import time
from pyrogram.errors.exceptions.bad_request_400 import UsernameNotOccupied
from pyrogram.errors import UsernameInvalid, FloodWait, UsernameNotOccupied
from pyrogram.raw import functions
from slugify import slugify
import openpyxl
from pyrogram import Client
from pyrogram.types import ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton, Message

main_group_id = 'seller_channel_haha'

file_ = openpyxl.load_workbook('list.xlsx')
sheet_obj_ = file_.active
categories = []
row_ = 1
file_categories = openpyxl.load_workbook('categories.xlsx')
sheet_obj_categories = file_categories.active

data_ = [({
    'channel': sheet_obj_.cell(row=i + 1, column=1).value,
    'category': sheet_obj_.cell(row=i + 1, column=2).value,
    'name_new': sheet_obj_.cell(row=i + 1, column=4).value,
    'id': sheet_obj_.cell(row=i + 1, column=5).value,
}) for i in range(1, sheet_obj_.max_row)]

for i in range(sheet_obj_categories.max_row-1):
    if sheet_obj_categories.cell(row=i + 1, column=1).value not in categories:
        categories.append(sheet_obj_categories.cell(row=i + 1, column=1).value)

file_categories.close()
file_.close()


def find_cell_by_link(v, sheet_obj):
    for row in sheet_obj.iter_rows():
        for cell in row:
            if cell.value == v:
                return row[0].row


async def main_create(app_bot, app_user):
    global data_
    for i in data_:
        print('fff')
        await asyncio.sleep(10)
        await create_channels(str(i.get('channel')), i, app_bot, app_user)


async def create_channel(username, title, data, channel, app_bot, app_user):
    await app_user.promote_chat_member(channel.id, "seller_test_s_bot")
    print(data.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_'))
    last_mes = None
    mg_id = 0
    async for message in app_user.get_chat_history(
            data.get('channel').replace('https://t.me/', '').replace('joinchat/', '').replace('-', '_'), limit=10):
        if not last_mes:
            last_mes = message
        if (message.photo or last_mes.photo or last_mes.media_group_id) and mg_id != message.media_group_id:
            last_mes = message
            kk = random.randrange(1, 4)
            await asyncio.sleep(kk)
            if message.media_group_id:
                try:
                    await app_user.copy_media_group(channel.id,
                                                    data.get('channel').replace('https://t.me/', '').replace(
                                                        'joinchat/', '').replace('-', '_'), message.id)
                except FloodWait as e:
                    print('wait for', e.value, 'to send message')
                    await asyncio.sleep(e.value + 2)
                    await app_user.copy_media_group(channel.id,
                                                    data.get('channel').replace('https://t.me/', '').replace(
                                                        'joinchat/', '').replace('-', '_'), message.id)
                except Exception as e:
                    print(e)
            elif message.photo:
                try:
                    await app_user.send_photo(channel.id, message.photo.file_id)
                except FloodWait as e:
                    print('wait for', e.value, 'to send message')
                    await asyncio.sleep(e.value + 2)
                    await app_user.send_photo(channel.id, message.photo.file_id)
                except Exception as e:
                    print(e)
            if message.text:
                try:
                    await app_user.send_message(channel.id, message.text)
                    print('text')
                except FloodWait as e:
                    print('wait for', e.value, 'to send message')
                    await asyncio.sleep(e.value + 2)
                    await app_user.send_message(channel.id, message.text)
                except Exception as e:
                    print(e)
            mg_id = message.media_group_id


async def create_channels(title, data, app_bot, app_user):
    username = "SBcr_" + slugify(title.replace('https://t.me/', '').replace('joinchat/', '')).replace('-', '_')
    if len(username) > 35:
        username = username[:34]
    file = openpyxl.load_workbook('list.xlsx')
    sheet_obj = file.active
    row = find_cell_by_link(title, sheet_obj)
    if sheet_obj.cell(row=row, column=5).value:
        try:
            msg = await app_user.get_messages(chat_id=sheet_obj.cell(row=row, column=5).value, message_ids=1)
        except Exception as e:
            print(e)
        file.close()
        file.close()
        await app_user.send_message('me', 'channel exists ' + str(sheet_obj.cell(row=row, column=4).value))
        print('channel exists')
    else:
        try:
            channel = await app_user.create_channel(sheet_obj.cell(row=row, column=4).value)
        except FloodWait as e:
            print('wait for', e.value, "to create channel")
            await app_user.send_message('me', 'wait for' + str(e.value) + "to create channel")
            await asyncio.sleep(e.value + 2)
            channel = await app_user.create_channel(sheet_obj.cell(row=row, column=4).value)
        finally:
            sheet_obj.cell(row=row, column=5, value=channel.id)
            file.save('list.xlsx')
            file.close()
            print(channel)
            await create_channel(username, title, data, channel, app_bot, app_user)


async def main_group_send_menu(app_bot, app_user):
    global main_group_id, data_
    CategoriesButtons = []
    for c in categories:
        GroupsButtons = []
        for g in [el for el in data_ if el.get('category') == c]:
            if g.get('id'):
                url = await app_user.create_chat_invite_link(g.get('id'))
                btn_g = InlineKeyboardButton(g.get('name_new'), url=url.invite_link)
                GroupsButtons.append([btn_g])
        await asyncio.sleep(1)
        if GroupsButtons:
            msg = await app_bot.send_message(main_group_id, c, reply_markup=InlineKeyboardMarkup(GroupsButtons))
        else:
            msg = await app_bot.send_message(main_group_id, c, reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton('Немає каналів', url='google.com')]]))
        file = openpyxl.load_workbook('categories.xlsx')
        sheet_obj = file.active
        row = find_cell_by_link(c, sheet_obj)
        print(row)
        sheet_obj.cell(row=row, column=2, value=msg.link)
        file.save('categories.xlsx')
        file.close()
        btn = InlineKeyboardButton(c, url=msg.link)
        CategoriesButtons.append([btn])
    CategoriesMarkup = InlineKeyboardMarkup(
        CategoriesButtons
    )

    print(main_group_id)
    await app_bot.send_message(main_group_id, 'Категорії', reply_markup=CategoriesMarkup)
#
# asyncio.run(main_group())
# asyncio.run(main())
