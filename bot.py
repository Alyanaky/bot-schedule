import telebot
import pandas as pd

bot = telebot.TeleBot('7832327916:AAFlm4hFis5qyz0emJchjUtYB_f4w_m0jWE')

df_lectures = pd.read_excel('k1l.xlsx')

days_lectures = {
 'Понедельник': 0,
 'Вторник': 1,
 'Среда': 2,
 'Четверг': 3,
 'Пятница': 4,
 'Суббота': 5
}

flow_files = {
 'В': 'k1sB.xlsx',
 'Г': 'k1sG.xlsx',
 'Д': 'k1sD.xlsx',
 'Е': 'k1sE.xlsx'
}

days_seminars = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']

keyboard_start = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
keyboard_start.add('Семинарские занятия', 'Лекционные занятия')

def get_streams_keyboard():
 keyboard = telebot.types.InlineKeyboardMarkup()
 keyboard.row(
  telebot.types.InlineKeyboardButton('В', callback_data='b'),
  telebot.types.InlineKeyboardButton('Г', callback_data='c'),
  telebot.types.InlineKeyboardButton('Д', callback_data='d'),
  telebot.types.InlineKeyboardButton('Е', callback_data='e')
 )
 return keyboard

def get_days_keyboard():
 keyboard = telebot.types.InlineKeyboardMarkup()
 for day in days_lectures.keys():
  keyboard.add(telebot.types.InlineKeyboardButton(day, callback_data=day))
 return keyboard

keyboard_start_seminars = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
keyboard_start_seminars.add('В', 'Г', 'Д', 'Е')

keyboard_group_seminars = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)

keyboard_day_seminars = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
keyboard_day_seminars.add(*days_seminars)

keyboard_back = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
keyboard_back.add('/get')

@bot.message_handler(commands=['start'])
def start(message):
 bot.send_message(message.chat.id, 'Привет! Я бот для получения расписания занятий. Выберите тип занятий:', reply_markup=keyboard_start)

@bot.message_handler(commands=['about'])
def about(message):
 bot.send_message(message.chat.id, 'Информация о боте:\n\nЭтот бот помогает студентам получать расписание лекционных и семинарских занятий.\n\nСсылка на разработчика: [Ваша ссылка]')

@bot.message_handler(commands=['get'])
def get_schedule(message):
 bot.send_message(message.chat.id, 'Выберите тип занятий:', reply_markup=keyboard_start)

@bot.message_handler(func=lambda message: message.text in ['Семинарские занятия', 'Лекционные занятия'])
def select_type(message):
 if message.text == 'Лекционные занятия':
  bot.send_message(message.chat.id, 'Выберите поток:', reply_markup=get_streams_keyboard())
 elif message.text == 'Семинарские занятия':
  bot.send_message(message.chat.id, 'Выберите поток:', reply_markup=keyboard_start_seminars)

@bot.callback_query_handler(func=lambda call: call.data in ['b', 'c', 'd', 'e'])
def handle_stream_choice(call):
  global stream
  stream = call.data
  bot.edit_message_text(
    chat_id=call.message.chat.id,
    message_id=call.message.message_id,
    text='Выберите день недели:',
    reply_markup=get_days_keyboard()
  )

@bot.callback_query_handler(func=lambda call: call.data in days_lectures.keys())
def handle_day_choice(call):
  day = call.data
  row = days_lectures[day]
  column = ord(stream) - ord('a')

  if row < len(df_lectures) and column < len(df_lectures.columns):
    if df_lectures.iloc[row, column] != 'Нет лекций':
      schedule = f'**{day}:**\n{df_lectures.iloc[row, column]}'
    else:
      schedule = 'Нет лекций'
  else:
    schedule = 'Ошибка: Неверный индекс'

  bot.edit_message_text(
    chat_id=call.message.chat.id,
    message_id=call.message.message_id,
    text=schedule
  )

@bot.message_handler(func=lambda message: message.text in flow_files.keys())
def select_flow(message):
  flow = message.text

  group_buttons = []
  if flow == 'В':
    group_buttons = list(range(132, 139))
  elif flow == 'Г':
    group_buttons = list(range(139, 146))
  elif flow == 'Д':
    group_buttons = list(range(146, 153))
  elif flow == 'Е':
    group_buttons = list(range(153, 161))

  keyboard_group_seminars.row(*[telebot.types.KeyboardButton(str(group)) for group in group_buttons])
  bot.send_message(message.chat.id, 'Выберите группу:', reply_markup=keyboard_group_seminars)
  bot.register_next_step_handler(message, select_group, flow)

@bot.message_handler(func=lambda message: message.text.isdigit())
def select_group(message, flow):
  global group
  group = int(message.text)
  bot.send_message(message.chat.id, 'Выберите день недели:', reply_markup=keyboard_day_seminars)
  bot.register_next_step_handler(message, select_day, flow)

@bot.message_handler(func=lambda message: message.text in days_seminars)
def select_day(message, flow):
  day = message.text
  schedule = get_schedule(flow, group, day)

  response = f"**{day}**\n**Группа {group}**\n\n"
  if schedule:
    response += f"*{schedule}*"
  else:
    response += "В этот день нет занятий."

  bot.send_message(message.chat.id, response, parse_mode='Markdown', reply_markup=keyboard_back)

def get_schedule(flow, group, day):
  df = pd.read_excel(flow_files[flow], sheet_name=0)
  try:
    day_index = days_seminars.index(day) + 2
    schedule = df.iloc[day_index, group - 132].strip()
  except:
    schedule = None
  return schedule

bot.polling(non_stop=True)