import zipfile
import xml.etree.ElementTree as ET
import random
import time
import urllib.request
import urllib.parse
import urllib.error
import json
import os
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

TOKEN = os.environ["BOT_TOKEN"]
XLSX  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "BT PEDAGOGIKASIDAN TESTLAR!!!.xlsx")

# {user_id: {'index': int, 'correct': int, 'wrong': int, 'skipped': int}}
user_state = {}


# ──────────────────────────────────────────────
#  Ma'lumotlarni yuklash
# ──────────────────────────────────────────────
def load_questions():
    with zipfile.ZipFile(XLSX) as z:
        shared_strings = []
        with z.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('ns:si', ns):
                texts = si.findall('.//ns:t', ns)
                shared_strings.append(''.join(t.text or '' for t in texts))
        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            rows = []
            for row in root.findall('.//ns:row', ns):
                row_data = []
                for cell in row.findall('ns:c', ns):
                    t = cell.get('t')
                    v_el = cell.find('ns:v', ns)
                    if v_el is not None:
                        val = v_el.text
                        if t == 's':
                            val = shared_strings[int(val)]
                    else:
                        val = ''
                    row_data.append(val)
                rows.append(row_data)

    questions = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        if i == 1:
            q       = row[1].strip() if len(row) > 1 else ''
            correct = row[2].strip() if len(row) > 2 else ''
            wrongs  = [row[j].strip() for j in range(3, min(6, len(row)))]
        else:
            q       = row[0].strip() if len(row) > 0 else ''
            correct = row[1].strip() if len(row) > 1 else ''
            wrongs  = [row[j].strip() for j in range(2, min(5, len(row)))]
        if not q or not correct:
            continue
        wrongs = [w for w in wrongs if w][:3]
        options = [correct] + wrongs
        random.seed(i * 17 + 3)
        random.shuffle(options)
        correct_idx = options.index(correct)
        q = q.replace('\n', ' ').strip()
        if len(q) > 300:
            q = q[:297] + '...'
        fixed_opts = []
        for opt in options:
            opt = opt.replace('\n', ' ').strip()
            if len(opt) > 100:
                opt = opt[:97] + '...'
            fixed_opts.append(opt)
        questions.append({'q': q, 'options': fixed_opts, 'ans': correct_idx})
    return questions


# ──────────────────────────────────────────────
#  Telegram API yordamchi funksiyalar
# ──────────────────────────────────────────────
def api(method, **params):
    url  = f"https://api.telegram.org/bot{TOKEN}/{method}"
    data = urllib.parse.urlencode(params).encode('utf-8')
    try:
        req = urllib.request.Request(url, data=data)
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"[API ERROR] {method}: {e.code} {body[:120]}")
        return {'ok': False, 'description': body}
    except Exception as e:
        print(f"[API ERROR] {method}: {e}")
        return {'ok': False}


def send_message(chat_id, text, reply_markup=None):
    params = dict(chat_id=chat_id, text=text, parse_mode='HTML')
    if reply_markup:
        params['reply_markup'] = json.dumps(reply_markup)
    api('sendMessage', **params)


def send_quiz_poll(chat_id, question, options, correct_option_id):
    return api(
        'sendPoll',
        chat_id=chat_id,
        question=question,
        options=json.dumps(options),
        type='quiz',
        correct_option_id=correct_option_id,
        is_anonymous=False,
    )


def answer_callback(callback_id, text=''):
    api('answerCallbackQuery', callback_query_id=callback_id, text=text)


# ──────────────────────────────────────────────
#  Inline tugmalar
# ──────────────────────────────────────────────
RESTART_KEYBOARD = {
    'inline_keyboard': [[
        {'text': '🔄 Qayta boshlash', 'callback_data': 'restart'},
        {'text': '📊 Natijani ko\'r',  'callback_data': 'result'},
    ]]
}

MENU_KEYBOARD = {
    'keyboard': [
        [{'text': '▶️ Boshlash'}, {'text': '📊 Natija'}],
        [{'text': '⏭ O\'tkazib yuborish'}, {'text': '⏹ To\'xtatish'}],
    ],
    'resize_keyboard': True,
}


# ──────────────────────────────────────────────
#  Savol yuborish
# ──────────────────────────────────────────────
def send_next_question(chat_id, questions):
    state = user_state.get(chat_id)
    if not state:
        return
    idx = state['index']

    if idx >= len(questions):
        finish_test(chat_id, questions)
        return

    item = questions[idx]
    num  = idx + 1
    q    = f"{num}/{len(questions)}. {item['q']}"
    if len(q) > 300:
        q = q[:297] + '...'

    send_quiz_poll(chat_id, q, item['options'], item['ans'])


def finish_test(chat_id, questions):
    state   = user_state.get(chat_id, {})
    correct = state.get('correct', 0)
    wrong   = state.get('wrong',   0)
    skipped = state.get('skipped', 0)
    total   = correct + wrong + skipped
    pct     = round(correct / total * 100) if total else 0

    if pct >= 90:
        emoji = '🏆'
        baho  = 'A\'lo!'
    elif pct >= 70:
        emoji = '🎉'
        baho  = 'Yaxshi!'
    elif pct >= 50:
        emoji = '👍'
        baho  = 'Qoniqarli'
    else:
        emoji = '📚'
        baho  = 'Ko\'proq o\'qing'

    send_message(
        chat_id,
        f"{emoji} <b>Test yakunlandi!</b>  —  {baho}\n\n"
        f"✅ To'g'ri javoblar:    <b>{correct}</b>\n"
        f"❌ Noto'g'ri javoblar: <b>{wrong}</b>\n"
        f"⏭ O'tkazib yuborildi: <b>{skipped}</b>\n"
        f"📊 Natija: <b>{pct}%</b>  ({correct}/{total})\n\n"
        f"Nima qilmoqchisiz?",
        reply_markup=RESTART_KEYBOARD
    )
    user_state.pop(chat_id, None)


def start_test(chat_id, name, questions):
    user_state[chat_id] = {'index': 0, 'correct': 0, 'wrong': 0, 'skipped': 0}
    send_message(
        chat_id,
        f"Salom, <b>{name}</b>! 👋\n\n"
        f"📚 <b>Boshlang'ich ta'lim pedagogikasi</b>\n"
        f"Jami: <b>{len(questions)}</b> ta savol\n\n"
        f"<b>Buyruqlar:</b>\n"
        f"/natija — hozirgi natija\n"
        f"/skip   — savolni o'tkazib yuborish\n"
        f"/stop   — testni to'xtatish\n"
        f"/help   — barcha buyruqlar\n\n"
        f"Boshlaylik! ⬇️",
        reply_markup=MENU_KEYBOARD
    )
    send_next_question(chat_id, questions)


# ──────────────────────────────────────────────
#  Update handler
# ──────────────────────────────────────────────
def handle_update(update, questions):

    # ── Matn xabarlari / buyruqlar ──
    if 'message' in update:
        msg     = update['message']
        chat_id = str(msg['chat']['id'])
        text    = msg.get('text', '').strip()
        name    = msg['from'].get('first_name', 'Do\'st')

        # /start yoki Boshlash tugmasi
        if text in ('/start', '▶️ Boshlash'):
            start_test(chat_id, name, questions)

        # /stop yoki To'xtatish
        elif text in ('/stop', '⏹ To\'xtatish'):
            if chat_id in user_state:
                state   = user_state.pop(chat_id)
                correct = state['correct']
                wrong   = state['wrong']
                total   = correct + wrong + state['skipped']
                pct     = round(correct / total * 100) if total else 0
                send_message(
                    chat_id,
                    f"⏹ <b>Test to'xtatildi.</b>\n\n"
                    f"✅ To'g'ri: <b>{correct}</b>\n"
                    f"❌ Noto'g'ri: <b>{wrong}</b>\n"
                    f"📊 Natija: <b>{pct}%</b>",
                    reply_markup=RESTART_KEYBOARD
                )
            else:
                send_message(chat_id, "Hozir aktiv test yo'q. Boshlash uchun /start yuboring.",
                             reply_markup=RESTART_KEYBOARD)

        # /natija yoki Natija tugmasi
        elif text in ('/natija', '📊 Natija'):
            if chat_id in user_state:
                s       = user_state[chat_id]
                total   = s['correct'] + s['wrong'] + s['skipped']
                pct     = round(s['correct'] / total * 100) if total else 0
                send_message(
                    chat_id,
                    f"📊 <b>Joriy natija:</b>\n\n"
                    f"📝 Savol: {s['index']}/{len(questions)}\n"
                    f"✅ To'g'ri: <b>{s['correct']}</b>\n"
                    f"❌ Noto'g'ri: <b>{s['wrong']}</b>\n"
                    f"⏭ O'tkazildi: <b>{s['skipped']}</b>\n"
                    f"📊 Foiz: <b>{pct}%</b>"
                )
            else:
                send_message(chat_id, "Hozir aktiv test yo'q. /start bilan boshlang.")

        # /skip yoki O'tkazib yuborish
        elif text in ('/skip', "⏭ O'tkazib yuborish"):
            if chat_id in user_state:
                user_state[chat_id]['skipped'] += 1
                user_state[chat_id]['index']   += 1
                send_next_question(chat_id, questions)
            else:
                send_message(chat_id, "Hozir aktiv test yo'q. /start bilan boshlang.")

        # /help
        elif text == '/help':
            send_message(
                chat_id,
                "📖 <b>Barcha buyruqlar:</b>\n\n"
                "/start  — testni boshlash (yoki qayta boshlash)\n"
                "/stop   — testni to'xtatish\n"
                "/natija — hozirgi natijani ko'rish\n"
                "/skip   — joriy savolni o'tkazib yuborish\n"
                "/help   — shu menyu\n\n"
                "💡 Pastdagi tugmalardan ham foydalanishingiz mumkin.",
                reply_markup=MENU_KEYBOARD
            )

        # Noma'lum xabar
        else:
            if chat_id not in user_state:
                send_message(
                    chat_id,
                    "Boshlash uchun /start yoki quyidagi tugmani bosing.",
                    reply_markup=RESTART_KEYBOARD
                )

    # ── Inline tugma bosildi ──
    elif 'callback_query' in update:
        cq      = update['callback_query']
        chat_id = str(cq['message']['chat']['id'])
        data    = cq.get('data', '')
        name    = cq['from'].get('first_name', 'Do\'st')

        if data == 'restart':
            answer_callback(cq['id'], '▶️ Qayta boshlanmoqda...')
            start_test(chat_id, name, questions)

        elif data == 'result':
            answer_callback(cq['id'])
            if chat_id in user_state:
                s     = user_state[chat_id]
                total = s['correct'] + s['wrong'] + s['skipped']
                pct   = round(s['correct'] / total * 100) if total else 0
                send_message(
                    chat_id,
                    f"📊 <b>Joriy natija:</b>\n"
                    f"✅ To'g'ri: <b>{s['correct']}</b>\n"
                    f"❌ Noto'g'ri: <b>{s['wrong']}</b>\n"
                    f"📊 Foiz: <b>{pct}%</b>"
                )
            else:
                send_message(chat_id, "Hozir aktiv test yo'q.", reply_markup=RESTART_KEYBOARD)

    # ── Quiz javobi ──
    elif 'poll_answer' in update:
        pa      = update['poll_answer']
        user_id = str(pa['user']['id'])

        if user_id not in user_state:
            return

        state = user_state[user_id]
        idx   = state['index']

        if pa['option_ids']:
            chosen = pa['option_ids'][0]
            if idx < len(questions) and chosen == questions[idx]['ans']:
                state['correct'] += 1
            else:
                state['wrong'] += 1
        else:
            state['skipped'] += 1

        state['index'] += 1
        time.sleep(0.3)
        send_next_question(user_id, questions)


# ──────────────────────────────────────────────
#  Asosiy tsikl
# ──────────────────────────────────────────────
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"OK")
    def log_message(self, *args):
        pass

def run_health_server():
    port = int(os.environ.get("PORT", 8080))
    HTTPServer(("0.0.0.0", port), HealthHandler).serve_forever()

def self_ping():
    url = os.environ.get("RENDER_EXTERNAL_URL", "")
    if not url:
        return
    while True:
        time.sleep(270)  # har 4.5 daqiqada
        try:
            urllib.request.urlopen(url, timeout=10)
        except Exception:
            pass


def main():
    print("=" * 45)
    print("  Savollar yuklanmoqda...")
    questions = load_questions()
    print(f"  Jami {len(questions)} ta savol yuklandi.")
    print(f"  Bot: @test_ishla_1bot")
    print(f"  To'xtatish: Ctrl+C")
    print("=" * 45)

    threading.Thread(target=run_health_server, daemon=True).start()
    threading.Thread(target=self_ping, daemon=True).start()
    print(f"  Health server: port {os.environ.get('PORT', 8080)}")

    offset = 0
    allowed = '["message","poll_answer","callback_query"]'

    while True:
        try:
            res = api('getUpdates', offset=offset, timeout=30, allowed_updates=allowed)
            if not res.get('ok'):
                time.sleep(3)
                continue
            for update in res.get('result', []):
                offset = update['update_id'] + 1
                try:
                    handle_update(update, questions)
                except Exception as e:
                    print(f"[HANDLER ERROR] {e}")
        except KeyboardInterrupt:
            print("\nBot to'xtatildi.")
            break
        except Exception as e:
            print(f"[LOOP ERROR] {e}")
            time.sleep(3)


if __name__ == '__main__':
    main()
