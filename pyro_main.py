import asyncio
import random
from pyrogram.errors import FloodWait
from slugify import slugify
import openpyxl
from pyrogram.types import InlineKeyboardMarkup, InlineKeyboardButton


def find_first_empty(self):
    r = 0
    while True:
        r += 1
        if not self.cell(r, 1).value:
            return r


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
}) for i in range(1, find_first_empty(sheet_obj_))]

for i in range(find_first_empty(sheet_obj_categories) - 1):
    if sheet_obj_categories.cell(row=i + 1, column=1).value not in categories:
        categories.append(sheet_obj_categories.cell(row=i + 1, column=1).value)

file_categories.close()
file_.close()


def find_cell_by_link(v, sheet_obj):
    i = 0
    while True:
        i += 1
        if sheet_obj.cell(i, 1).value == v:
            return i


async def main_create(app_bot, app_user, CallbackQuery):
    global data_
    for i in data_:
        print('iter')
        await asyncio.sleep(10)
        await create_channels(str(i.get('channel')), i, app_bot, app_user, CallbackQuery)
    await CallbackQuery.answer()


async def create_channel(data, channel, app_user):
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


async def create_channels(title, data, app_bot, app_user, CallbackQuery):
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
        await app_bot.send_message(CallbackQuery.message.chat.id,
                                   'Канал вже існує: ' + str(sheet_obj.cell(row=row, column=4).value))
        print('channel exists')
    else:
        try:
            channel = await app_user.create_channel(sheet_obj.cell(row=row, column=4).value)
        except FloodWait as e:
            print('wait for', e.value, "to create channel")
            await app_bot.send_message(CallbackQuery.message.chat.id,
                                       'Почекайте ' + str(e.value) + " секунди щоб створити канал")
            await asyncio.sleep(e.value + 2)
            channel = await app_user.create_channel(sheet_obj.cell(row=row, column=4).value)
        finally:
            sheet_obj.cell(row=row, column=5, value=channel.id)
            file.save('list.xlsx')
            file.close()
            await asyncio.sleep(3)
            await create_channel(data, channel, app_user)


async def main_group_send_menu(app_bot, app_user, CallbackQuery):
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
            file = openpyxl.load_workbook('categories.xlsx')
            sheet_obj = file.active
            row = find_cell_by_link(c, sheet_obj)

            sheet_obj.cell(row=row, column=2, value=msg.link)
            file.save('categories.xlsx')
            file.close()
            print(sheet_obj.cell(row=row, column=2).value)
            btn = InlineKeyboardButton(c, url=msg.link)
            CategoriesButtons.append([btn])
    CategoriesMarkup = InlineKeyboardMarkup(
        CategoriesButtons
    )

    cat_m = await app_bot.send_message(main_group_id, 'Категорії', reply_markup=CategoriesMarkup)
    file = openpyxl.load_workbook('categories.xlsx')
    sheet_obj = file.active
    row = find_cell_by_link('Категорії', sheet_obj)
    sheet_obj.cell(row=row, column=2, value=cat_m.link)
    file.save('categories.xlsx')
    file.close()
    await CallbackQuery.answer()


async def send_new_msg(app_user, message, new_channel_id, CallbackQuery):
    last_mes = await app_user.get_messages(message.chat.id, message.id - 1)
    mg_id = await app_user.get_messages(message.chat.id, message.id - 1).media_group_id
    if (message.photo or last_mes.photo or last_mes.media_group_id) and mg_id != message.media_group_id:
        kk = random.randrange(1, 4)
        await asyncio.sleep(kk)
        if message.media_group_id:
            try:
                await app_user.copy_media_group(new_channel_id,
                                                CallbackQuery.message.chat.id, message.id)
            except FloodWait as e:
                print('wait for', e.value, 'to send message')
                await asyncio.sleep(e.value + 2)
                await app_user.copy_media_group(new_channel_id,
                                                CallbackQuery.message.chat.id, message.id)
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
    await CallbackQuery.answer()


async def add_new_category(app_bot, message):
    file_c = openpyxl.load_workbook('categories.xlsx')
    sheet_obj_n = file_c.active
    row = find_first_empty(sheet_obj_n)
    name = message.text.replace('/add_category', '').strip()
    categories = []
    for i in range(find_first_empty(sheet_obj_n)):
        if sheet_obj_n.cell(row=i + 1, column=1).value not in categories:
            categories.append(sheet_obj_n.cell(row=i + 1, column=1).value)
    if name not in categories:
        sheet_obj_n.cell(row, 1, value=name)
        file_c.save('categories.xlsx')
        file_c.close()
        await app_bot.send_message(message.chat.id, 'Категорія додана')
    else:
        await app_bot.send_message(message.chat.id, 'Така категорія вже існує')
    await asyncio.sleep(3)


async def add_new_channel(app_bot, message):
    file_c = openpyxl.load_workbook('list.xlsx')
    sheet_obj_n = file_c.active
    row = find_first_empty(sheet_obj_n)
    name = message.text.replace('/add_channel', '')
    channels = []
    for i in range(1, find_first_empty(sheet_obj_n)):
        if sheet_obj_n.cell(row=i + 1, column=1).value not in channels:
            channels.append(sheet_obj_n.cell(row=i + 1, column=1).value)
    name = name.split(',')
    if name[0] not in channels:
        try:
            sheet_obj_n.cell(row, 1, value=name[0])
            sheet_obj_n.cell(row, 2, value=name[1])
            sheet_obj_n.cell(row, 4, value=name[2])
            await app_bot.send_message(message.chat.id, 'Канал додано')

            file_c.save('list.xlsx')
        except Exception as e:
            print(e)
            await app_bot.send_message(message.chat.id, 'Неправильно введені дані')


    else:
        await app_bot.send_message(message.chat.id, 'Такий канал вже додано')
    file_c.close()
    await asyncio.sleep(3)


async def delete_channel(app_bot, message, CallbackQuery):
    file_c = openpyxl.load_workbook('list.xlsx')
    sheet_obj_n = file_c.active
    # row = find_first_empty(sheet_obj_n)
    name = message.replace('/delete_channel', '')
    channels = []
    for i in range(1, find_first_empty(sheet_obj_n)):
        if sheet_obj_n.cell(row=i + 1, column=1).value not in channels:
            channels.append(sheet_obj_n.cell(row=i + 1, column=1).value)
    name = name.replace(' ', '')
    print(name)
    if name in channels:
        try:
            row = find_cell_by_link(name, sheet_obj_n)
            sheet_obj_n.delete_rows(row, 1)
        except Exception as e:
            print(e)
            await app_bot.send_message(CallbackQuery.message.chat.id, 'Помилка при видаленні')

        file_c.save('list.xlsx')
        file_c.close()
        await app_bot.send_message(CallbackQuery.message.chat.id, 'Канал видалено')
    else:
        await app_bot.send_message(CallbackQuery.message.chat.id, 'Немає такого каналу')
    await CallbackQuery.answer()
    await asyncio.sleep(3)


async def delete_category(app_bot, message, CallbackQuery):
    file_c = openpyxl.load_workbook('categories.xlsx')
    sheet_obj_n = file_c.active
    # row = find_first_empty(sheet_obj_n)
    name = message.replace('/delete_category', '').strip()
    categories = []
    for i in range(find_first_empty(sheet_obj_n)):
        if sheet_obj_n.cell(row=i + 1, column=1).value not in categories:
            categories.append(sheet_obj_n.cell(row=i + 1, column=1).value)
    name = name
    print(name)
    deleted = False
    if name in categories:
        try:
            row = find_cell_by_link(name, sheet_obj_n)
            sheet_obj_n.delete_rows(row, 1)
            deleted = True
        except Exception as e:
            print(e)
            await app_bot.send_message(CallbackQuery.message.chat.id, 'Помилка при видаленні')
        file_c.save('categories.xlsx')
        file_c.close()

        await app_bot.send_message(CallbackQuery.message.chat.id, 'Категорія видалена')
        print(deleted)
        if deleted:
            file_c = openpyxl.load_workbook('list.xlsx')
            sheet_obj_n = file_c.active
            for i in range(1, find_first_empty(sheet_obj_n)):
                if sheet_obj_n.cell(i, 2).value == name:
                    sheet_obj_n.delete_rows(i, 1)
                    file_c.save('list.xlsx')
            file_c.close()
    else:
        await app_bot.send_message(CallbackQuery.message.chat.id, 'Немає такої категорії')
    await CallbackQuery.answer()
    await asyncio.sleep(3)


async def main_group_delete_menu(app_user, CallbackQuery):
    file_c = openpyxl.load_workbook('categories.xlsx')
    sheet_obj_n = file_c.active
    menu_msgs = []
    for i in range(1, find_first_empty(sheet_obj_n)):
        if sheet_obj_n.cell(i, 2).value:
            menu_msgs.append(int(sheet_obj_n.cell(i, 2).value.split('/')[-1]))
    await app_user.delete_messages(main_group_id, menu_msgs)
    await CallbackQuery.answer()
    await asyncio.sleep(3)
