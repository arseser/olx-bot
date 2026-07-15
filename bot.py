import os
import json
import smtplib
import dns.resolver
import pandas as pd
import time
import telebot
from telebot import types

TOKEN = "8921710608:AAHyvGTupv9uutrfKf4lWCblA-pUKFLXm8I"
bot = telebot.TeleBot(TOKEN)

DOMAINS = ["@wp.pl", "@vp.pl"]
PAUSE_BETWEEN_CHECKS = 1.0
MAX_RETRIES = 2

# ==========================================
# ПРОВЕРКА EMAIL ЧЕРЕЗ SMTP
# ==========================================
def get_mx_servers(domain):
    try:
        answers = dns.resolver.resolve(domain, 'MX')
        mx_list = []
        for r in answers:
            mx_list.append((r.preference, str(r.exchange).rstrip('.')))
        mx_list.sort(key=lambda x: x[0])
        return mx_list
    except Exception:
        return []

def check_one_email(email):
    try:
        domain = email.split('@')[1]
        mx_servers = get_mx_servers(domain)
        
        if not mx_servers:
            return False
        
        for pref, mx_server in mx_servers:
            for attempt in range(MAX_RETRIES):
                try:
                    with smtplib.SMTP(mx_server, 25, timeout=15) as smtp:
                        smtp.helo('olx-checker')
                        smtp.mail_from('check@example.com')
                        code, message = smtp.rcpt(email)
                        
                        if code == 250 or code == 251:
                            return True
                        elif code == 550:
                            return False
                        else:
                            break
                except Exception:
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(1)
                    continue
            continue
        
        return False
    except Exception:
        return False

# ==========================================
# ОБРАБОТКА JSON ФАЙЛА
# ==========================================
def process_olx_data(data, progress_callback=None):
    results = []
    total = len(data)
    checked = 0
    skipped = 0
    
    for key, ad in data.items():
        seller = ""
        should_check = False
        
        if "1seller" in ad:
            seller = ad["1seller"].strip()
            should_check = True
        elif "seller" in ad:
            seller = ad["seller"].strip()
            should_check = False
        
        ad_url = ad.get("ad_url", "").strip()
        title = ad.get("title", "").strip()
        
        if not seller or not ad_url:
            continue
        
        if not should_check:
            skipped += 1
            checked += 1
            continue
        
        wp_mail = f"{seller}@wp.pl"
        vp_mail = f"{seller}@vp.pl"
        
        time.sleep(PAUSE_BETWEEN_CHECKS)
        wp_valid = check_one_email(wp_mail)
        
        time.sleep(PAUSE_BETWEEN_CHECKS)
        vp_valid = check_one_email(vp_mail)
        
        note_parts = []
        valid_emails = []
        
        if wp_valid:
            valid_emails.append(wp_mail)
            note_parts.append("WP: ✅")
        else:
            note_parts.append("WP: ❌")
        
        if vp_valid:
            valid_emails.append(vp_mail)
            note_parts.append("VP: ✅")
        else:
            note_parts.append("VP: ❌")
        
        note = " | ".join(note_parts)
        
        if valid_emails:
            for email in valid_emails:
                results.append({
                    "Email": email,
                    "Ссылка на объявление": ad_url,
                    "Название объявления": title,
                    "Продавец (seller)": seller,
                    "Результат проверки": note
                })
        
        checked += 1
    
    return results, skipped

# ==========================================
# КОМАНДЫ БОТА
# ==========================================
@bot.message_handler(commands=['start'])
def send_welcome(message):
    welcome_text = (
        "📬 *OLX Email Checker Bot*\n\n"
        "Привет! Я проверяю почты @wp.pl и @vp.pl из объявлений OLX.\n\n"
        "🔹 *Как пользоваться:*\n"
        "1. Отправь мне TXT-файл с объявлениями (JSON)\n"
        "2. Я проверю всех у кого поле `1seller`\n"
        "3. Пришлю тебе готовый Excel с валидными почтами\n\n"
        "⚠️ Поле `seller` (без 1) я пропускаю — проверяю только `1seller`!"
    )
    bot.reply_to(message, welcome_text, parse_mode='Markdown')

@bot.message_handler(content_types=['document'])
def handle_file(message):
    try:
        # Проверяем расширение
        file_name = message.document.file_name
        if not file_name.endswith('.txt'):
            bot.reply_to(message, "❌ Пришли TXT-файл, а не что-то другое!")
            return
        
        # Скачиваем файл
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Парсим JSON
        data = json.loads(downloaded_file.decode('utf-8'))
        
        total_ads = len(data)
        
        # Отправляем сообщение о начале работы
        status_msg = bot.reply_to(message, f"⏳ Начинаю проверку...\nВсего объявлений: {total_ads}\nЭто может занять несколько минут.")
        
        # Обрабатываем
        results, skipped = process_olx_data(data)
        
        # Создаём Excel
        output_file = "результат_проверки.xlsx"
        if results:
            df = pd.DataFrame(results)
            df.to_excel(output_file, index=False, engine='openpyxl')
        else:
            df = pd.DataFrame(columns=["Email", "Ссылка на объявление", "Название объявления", "Продавец (seller)", "Результат проверки"])
            df.to_excel(output_file, index=False, engine='openpyxl')
        
        # Отправляем результат
        with open(output_file, 'rb') as f:
            caption = (
                f"✅ *Готово!*\n\n"
                f"📊 Всего объявлений: {total_ads}\n"
                f"⏭ Пропущено (seller): {skipped}\n"
                f"🔍 Проверено (1seller): {total_ads - skipped}\n"
                f"📬 Найдено валидных: {len(results)}"
            )
            bot.send_document(message.chat.id, f, caption=caption, parse_mode='Markdown')
        
        # Удаляем временный файл
        os.remove(output_file)
        
        # Обновляем статус
        bot.edit_message_text("✅ Проверка завершена! Результат выше 👆", message.chat.id, status_msg.message_id)
        
    except json.JSONDecodeError:
        bot.reply_to(message, "❌ Ошибка: файл не в формате JSON! Проверь что это правильный TXT с объявлениями.")
    except Exception as e:
        bot.reply_to(message, f"❌ Произошла ошибка:\n`{str(e)}`", parse_mode='Markdown')

@bot.message_handler(func=lambda message: True)
def handle_other(message):
    bot.reply_to(message, "👋 Пришли мне TXT-файл с объявлениями, и я проверю почты!\n\nПодробнее: /start")

# ==========================================
# ЗАПУСК БОТА
# ==========================================
# Удаляем webhook если он был
bot.remove_webhook()
time.sleep(1)
print("=" * 50)
print("  OLX EMAIL CHECKER BOT — запускаюсь!")
print("=" * 50)

bot.infinity_polling()
